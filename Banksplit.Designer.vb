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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Banksplit))
        Me.Label76 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Dgv_Split = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Btn_Split_Save = New System.Windows.Forms.Button()
        Me.Btn_Split_Cancel = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Lbl_SplitBank_Accountnr = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Lbl_SplitBank_Type = New System.Windows.Forms.Label()
        Me.Btn_Prefill_Split = New System.Windows.Forms.Button()
        Me.Lbl_Split_Description = New System.Windows.Forms.Label()
        Me.Lbl_Split_Diff = New System.Windows.Forms.Label()
        Me.Lbl_Split_Amount = New System.Windows.Forms.Label()
        Me.Lbl_Split_Bank_id = New System.Windows.Forms.Label()
        Me.Lbl_Split_seqorder = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Lbl_SplitBank_journal_name = New System.Windows.Forms.Label()
        CType(Me.Dgv_Split, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label76
        '
        Me.Label76.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label76.Location = New System.Drawing.Point(3, 9)
        Me.Label76.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label76.Name = "Label76"
        Me.Label76.Size = New System.Drawing.Size(91, 19)
        Me.Label76.TabIndex = 31
        Me.Label76.Text = "Omschrijving"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 90)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(55, 19)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Bedrag"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(287, 9)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(74, 19)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "Bank /volgnr"
        '
        'Dgv_Split
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Dgv_Split.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Blue
        DataGridViewCellStyle2.NullValue = "0"
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.Dgv_Split.DefaultCellStyle = DataGridViewCellStyle2
        Me.Dgv_Split.Location = New System.Drawing.Point(15, 113)
        Me.Dgv_Split.Name = "Dgv_Split"
        Me.Dgv_Split.Size = New System.Drawing.Size(502, 203)
        Me.Dgv_Split.TabIndex = 35
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(173, 88)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 19)
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
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(288, 28)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(60, 19)
        Me.Label5.TabIndex = 119
        Me.Label5.Text = "Account"
        '
        'Lbl_SplitBank_Accountnr
        '
        Me.Lbl_SplitBank_Accountnr.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_SplitBank_Accountnr.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_SplitBank_Accountnr.Location = New System.Drawing.Point(366, 28)
        Me.Lbl_SplitBank_Accountnr.Name = "Lbl_SplitBank_Accountnr"
        Me.Lbl_SplitBank_Accountnr.Size = New System.Drawing.Size(151, 19)
        Me.Lbl_SplitBank_Accountnr.TabIndex = 120
        Me.Lbl_SplitBank_Accountnr.Text = "Accountnr"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(288, 49)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(51, 19)
        Me.Label6.TabIndex = 121
        Me.Label6.Text = "Type"
        '
        'Lbl_SplitBank_Type
        '
        Me.Lbl_SplitBank_Type.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_SplitBank_Type.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_SplitBank_Type.Location = New System.Drawing.Point(366, 49)
        Me.Lbl_SplitBank_Type.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Lbl_SplitBank_Type.Name = "Lbl_SplitBank_Type"
        Me.Lbl_SplitBank_Type.Size = New System.Drawing.Size(72, 19)
        Me.Lbl_SplitBank_Type.TabIndex = 122
        Me.Lbl_SplitBank_Type.Text = "Type"
        '
        'Btn_Prefill_Split
        '
        Me.Btn_Prefill_Split.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Prefill_Split.Location = New System.Drawing.Point(14, 329)
        Me.Btn_Prefill_Split.Name = "Btn_Prefill_Split"
        Me.Btn_Prefill_Split.Size = New System.Drawing.Size(318, 29)
        Me.Btn_Prefill_Split.TabIndex = 123
        Me.Btn_Prefill_Split.Text = "Vullen met openstaande uitkeringsformulieren"
        Me.Btn_Prefill_Split.UseVisualStyleBackColor = True
        '
        'Lbl_Split_Description
        '
        Me.Lbl_Split_Description.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_Split_Description.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Lbl_Split_Description.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Split_Description.Location = New System.Drawing.Point(91, 9)
        Me.Lbl_Split_Description.Name = "Lbl_Split_Description"
        Me.Lbl_Split_Description.Size = New System.Drawing.Size(191, 38)
        Me.Lbl_Split_Description.TabIndex = 124
        Me.Lbl_Split_Description.Text = "Label7"
        '
        'Lbl_Split_Diff
        '
        Me.Lbl_Split_Diff.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_Split_Diff.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Lbl_Split_Diff.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Split_Diff.Location = New System.Drawing.Point(230, 88)
        Me.Lbl_Split_Diff.Name = "Lbl_Split_Diff"
        Me.Lbl_Split_Diff.Size = New System.Drawing.Size(79, 22)
        Me.Lbl_Split_Diff.TabIndex = 125
        Me.Lbl_Split_Diff.Text = "0.00"
        Me.Lbl_Split_Diff.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Lbl_Split_Amount
        '
        Me.Lbl_Split_Amount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_Split_Amount.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Lbl_Split_Amount.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Split_Amount.Location = New System.Drawing.Point(82, 88)
        Me.Lbl_Split_Amount.Name = "Lbl_Split_Amount"
        Me.Lbl_Split_Amount.Size = New System.Drawing.Size(86, 22)
        Me.Lbl_Split_Amount.TabIndex = 126
        Me.Lbl_Split_Amount.Text = "Label7"
        Me.Lbl_Split_Amount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Lbl_Split_Bank_id
        '
        Me.Lbl_Split_Bank_id.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_Split_Bank_id.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Split_Bank_id.Location = New System.Drawing.Point(366, 7)
        Me.Lbl_Split_Bank_id.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Lbl_Split_Bank_id.Name = "Lbl_Split_Bank_id"
        Me.Lbl_Split_Bank_id.Size = New System.Drawing.Size(49, 19)
        Me.Lbl_Split_Bank_id.TabIndex = 127
        Me.Lbl_Split_Bank_id.Text = "id"
        '
        'Lbl_Split_seqorder
        '
        Me.Lbl_Split_seqorder.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_Split_seqorder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Split_seqorder.Location = New System.Drawing.Point(420, 7)
        Me.Lbl_Split_seqorder.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Lbl_Split_seqorder.Name = "Lbl_Split_seqorder"
        Me.Lbl_Split_seqorder.Size = New System.Drawing.Size(68, 19)
        Me.Lbl_Split_seqorder.TabIndex = 128
        Me.Lbl_Split_seqorder.Text = "seq"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 47)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 19)
        Me.Label4.TabIndex = 129
        Me.Label4.Text = "Journaal ref"
        '
        'Lbl_SplitBank_journal_name
        '
        Me.Lbl_SplitBank_journal_name.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Lbl_SplitBank_journal_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_SplitBank_journal_name.Location = New System.Drawing.Point(91, 49)
        Me.Lbl_SplitBank_journal_name.Name = "Lbl_SplitBank_journal_name"
        Me.Lbl_SplitBank_journal_name.Size = New System.Drawing.Size(191, 19)
        Me.Lbl_SplitBank_journal_name.TabIndex = 130
        Me.Lbl_SplitBank_journal_name.Text = "fk_journal_name"
        '
        'Banksplit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(526, 370)
        Me.Controls.Add(Me.Lbl_SplitBank_journal_name)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Lbl_Split_seqorder)
        Me.Controls.Add(Me.Lbl_Split_Bank_id)
        Me.Controls.Add(Me.Lbl_Split_Amount)
        Me.Controls.Add(Me.Lbl_Split_Diff)
        Me.Controls.Add(Me.Lbl_Split_Description)
        Me.Controls.Add(Me.Btn_Prefill_Split)
        Me.Controls.Add(Me.Lbl_SplitBank_Type)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Lbl_SplitBank_Accountnr)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Btn_Split_Cancel)
        Me.Controls.Add(Me.Btn_Split_Save)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Dgv_Split)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label76)
        Me.Name = "Banksplit"
        Me.Text = "Banktransacties splitsen"
        CType(Me.Dgv_Split, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label76 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Dgv_Split As DataGridView
    Friend WithEvents Label3 As Label
    Friend WithEvents Btn_Split_Save As Button
    Friend WithEvents Btn_Split_Cancel As Button
    Friend WithEvents Label5 As Label
    Friend WithEvents Lbl_SplitBank_Accountnr As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Lbl_SplitBank_Type As Label
    Friend WithEvents Btn_Prefill_Split As Button
    Friend WithEvents Lbl_Split_Description As Label
    Friend WithEvents Lbl_Split_Diff As Label
    Friend WithEvents Lbl_Split_Amount As Label
    Friend WithEvents Lbl_Split_Bank_id As Label
    Friend WithEvents Lbl_Split_seqorder As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Lbl_SplitBank_journal_name As Label
End Class
