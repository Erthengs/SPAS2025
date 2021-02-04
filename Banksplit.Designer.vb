<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Banksplit
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Banksplit))
        Me.Label76 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Tbx_Split_Description = New System.Windows.Forms.TextBox()
        Me.Tbx_Split_Amount = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Tbx_Split_seqorder = New System.Windows.Forms.TextBox()
        Me.Dgv_Split = New System.Windows.Forms.DataGridView()
        Me.Tbx_Split_Diff = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Btn_Split_Save = New System.Windows.Forms.Button()
        Me.Btn_Split_Cancel = New System.Windows.Forms.Button()
        Me.Tbx_Split_Bank_id = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.Dgv_Split, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label76
        '
        Me.Label76.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label76.Location = New System.Drawing.Point(24, 28)
        Me.Label76.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(138, 18)
        Me.Label76.TabIndex = 31
        Me.Label76.Text = "Omschrijving"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 86)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 19)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Bedrag"
        '
        'Tbx_Split_Description
        '
        Me.Tbx_Split_Description.Enabled = False
        Me.Tbx_Split_Description.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbx_Split_Description.Location = New System.Drawing.Point(142, 28)
        Me.Tbx_Split_Description.Multiline = True
        Me.Tbx_Split_Description.Name = "Tbx_Split_Description"
        Me.Tbx_Split_Description.Size = New System.Drawing.Size(373, 51)
        Me.Tbx_Split_Description.TabIndex = 33
        '
        'Tbx_Split_Amount
        '
        Me.Tbx_Split_Amount.Enabled = False
        Me.Tbx_Split_Amount.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbx_Split_Amount.Location = New System.Drawing.Point(142, 85)
        Me.Tbx_Split_Amount.Name = "Tbx_Split_Amount"
        Me.Tbx_Split_Amount.Size = New System.Drawing.Size(100, 21)
        Me.Tbx_Split_Amount.TabIndex = 34
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(287, 88)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(113, 19)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "Bank id"
        '
        'Tbx_Split_seqorder
        '
        Me.Tbx_Split_seqorder.Enabled = False
        Me.Tbx_Split_seqorder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbx_Split_seqorder.Location = New System.Drawing.Point(414, 110)
        Me.Tbx_Split_seqorder.Name = "Tbx_Split_seqorder"
        Me.Tbx_Split_seqorder.Size = New System.Drawing.Size(100, 21)
        Me.Tbx_Split_seqorder.TabIndex = 34
        '
        'Dgv_Split
        '
        Me.Dgv_Split.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.Dgv_Split.Location = New System.Drawing.Point(26, 142)
        Me.Dgv_Split.Name = "Dgv_Split"
        Me.Dgv_Split.Size = New System.Drawing.Size(488, 174)
        Me.Dgv_Split.TabIndex = 35
        '
        'Tbx_Split_Diff
        '
        Me.Tbx_Split_Diff.Enabled = False
        Me.Tbx_Split_Diff.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbx_Split_Diff.Location = New System.Drawing.Point(142, 115)
        Me.Tbx_Split_Diff.Name = "Tbx_Split_Diff"
        Me.Tbx_Split_Diff.Size = New System.Drawing.Size(100, 21)
        Me.Tbx_Split_Diff.TabIndex = 37
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(24, 116)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 19)
        Me.Label3.TabIndex = 36
        Me.Label3.Text = "Verschil"
        '
        'Btn_Split_Save
        '
        Me.Btn_Split_Save.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Btn_Split_Save.Image = CType(resources.GetObject("Btn_Split_Save.Image"), System.Drawing.Image)
        Me.Btn_Split_Save.Location = New System.Drawing.Point(479, 322)
        Me.Btn_Split_Save.Name = "Btn_Split_Save"
        Me.Btn_Split_Save.Size = New System.Drawing.Size(35, 36)
        Me.Btn_Split_Save.TabIndex = 97
        Me.Btn_Split_Save.UseVisualStyleBackColor = True
        '
        'Btn_Split_Cancel
        '
        Me.Btn_Split_Cancel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Btn_Split_Cancel.Image = Global.SPAS.My.Resources.Resources.Cancel_16x16
        Me.Btn_Split_Cancel.Location = New System.Drawing.Point(438, 322)
        Me.Btn_Split_Cancel.Name = "Btn_Split_Cancel"
        Me.Btn_Split_Cancel.Size = New System.Drawing.Size(35, 36)
        Me.Btn_Split_Cancel.TabIndex = 116
        Me.Btn_Split_Cancel.UseVisualStyleBackColor = True
        '
        'Tbx_Split_Bank_id
        '
        Me.Tbx_Split_Bank_id.Enabled = False
        Me.Tbx_Split_Bank_id.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbx_Split_Bank_id.Location = New System.Drawing.Point(414, 85)
        Me.Tbx_Split_Bank_id.Name = "Tbx_Split_Bank_id"
        Me.Tbx_Split_Bank_id.Size = New System.Drawing.Size(100, 21)
        Me.Tbx_Split_Bank_id.TabIndex = 117
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(287, 113)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(113, 19)
        Me.Label4.TabIndex = 118
        Me.Label4.Text = "Volgnummer"
        '
        'Banksplit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(526, 370)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Tbx_Split_Bank_id)
        Me.Controls.Add(Me.Btn_Split_Cancel)
        Me.Controls.Add(Me.Btn_Split_Save)
        Me.Controls.Add(Me.Tbx_Split_Diff)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Dgv_Split)
        Me.Controls.Add(Me.Tbx_Split_seqorder)
        Me.Controls.Add(Me.Tbx_Split_Amount)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Tbx_Split_Description)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label76)
        Me.Name = "Banksplit"
        Me.Text = "Banktransacties splitsen"
        CType(Me.Dgv_Split, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label76 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Tbx_Split_Description As TextBox
    Friend WithEvents Tbx_Split_Amount As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Tbx_Split_seqorder As TextBox
    Friend WithEvents Dgv_Split As DataGridView
    Friend WithEvents Tbx_Split_Diff As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Btn_Split_Save As Button
    Friend WithEvents Btn_Split_Cancel As Button
    Friend WithEvents Tbx_Split_Bank_id As TextBox
    Friend WithEvents Label4 As Label
End Class
