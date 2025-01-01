<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Helptext
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Helptext))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.LbL_Onderwerp = New System.Windows.Forms.Label()
        Me.LbL_Onderwerp2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(12, 35)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(453, 320)
        Me.TextBox1.TabIndex = 1
        '
        'LbL_Onderwerp
        '
        Me.LbL_Onderwerp.AutoSize = True
        Me.LbL_Onderwerp.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbL_Onderwerp.Location = New System.Drawing.Point(14, 14)
        Me.LbL_Onderwerp.Name = "LbL_Onderwerp"
        Me.LbL_Onderwerp.Size = New System.Drawing.Size(85, 18)
        Me.LbL_Onderwerp.TabIndex = 2
        Me.LbL_Onderwerp.Text = "Onderwerp:"
        '
        'LbL_Onderwerp2
        '
        Me.LbL_Onderwerp2.AutoSize = True
        Me.LbL_Onderwerp2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LbL_Onderwerp2.Location = New System.Drawing.Point(96, 14)
        Me.LbL_Onderwerp2.Name = "LbL_Onderwerp2"
        Me.LbL_Onderwerp2.Size = New System.Drawing.Size(20, 18)
        Me.LbL_Onderwerp2.TabIndex = 3
        Me.LbL_Onderwerp2.Text = "..."
        '
        'Helptext
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(477, 367)
        Me.Controls.Add(Me.LbL_Onderwerp2)
        Me.Controls.Add(Me.LbL_Onderwerp)
        Me.Controls.Add(Me.TextBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Helptext"
        Me.Text = "Ondersteuning"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents LbL_Onderwerp As Label
    Friend WithEvents LbL_Onderwerp2 As Label
End Class
