<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.ImportBtn = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.StatusLbl = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.VoucherDateField = New System.Windows.Forms.DateTimePicker()
        Me.TransTempltBtn = New System.Windows.Forms.Button()
        Me.TranImportBtn = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.DividendTempltBtn = New System.Windows.Forms.Button()
        Me.ImportDividendBtn = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImportBtn
        '
        Me.ImportBtn.Enabled = False
        Me.ImportBtn.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ImportBtn.Location = New System.Drawing.Point(17, 19)
        Me.ImportBtn.Name = "ImportBtn"
        Me.ImportBtn.Size = New System.Drawing.Size(109, 23)
        Me.ImportBtn.TabIndex = 0
        Me.ImportBtn.Text = "Master Import"
        Me.ImportBtn.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ImportBtn)
        Me.GroupBox1.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(4, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(447, 52)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Masters"
        '
        'StatusLbl
        '
        Me.StatusLbl.BackColor = System.Drawing.Color.White
        Me.StatusLbl.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusLbl.Location = New System.Drawing.Point(1, 193)
        Me.StatusLbl.Name = "StatusLbl"
        Me.StatusLbl.Size = New System.Drawing.Size(450, 18)
        Me.StatusLbl.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.VoucherDateField)
        Me.GroupBox2.Controls.Add(Me.TransTempltBtn)
        Me.GroupBox2.Controls.Add(Me.TranImportBtn)
        Me.GroupBox2.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(4, 61)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(447, 56)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Transaction"
        '
        'VoucherDateField
        '
        Me.VoucherDateField.CalendarFont = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VoucherDateField.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.VoucherDateField.Location = New System.Drawing.Point(17, 21)
        Me.VoucherDateField.Name = "VoucherDateField"
        Me.VoucherDateField.Size = New System.Drawing.Size(109, 20)
        Me.VoucherDateField.TabIndex = 2
        '
        'TransTempltBtn
        '
        Me.TransTempltBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TransTempltBtn.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TransTempltBtn.Location = New System.Drawing.Point(316, 19)
        Me.TransTempltBtn.Name = "TransTempltBtn"
        Me.TransTempltBtn.Size = New System.Drawing.Size(124, 24)
        Me.TransTempltBtn.TabIndex = 1
        Me.TransTempltBtn.Text = "Template"
        Me.TransTempltBtn.UseVisualStyleBackColor = False
        '
        'TranImportBtn
        '
        Me.TranImportBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.TranImportBtn.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TranImportBtn.Location = New System.Drawing.Point(136, 19)
        Me.TranImportBtn.Name = "TranImportBtn"
        Me.TranImportBtn.Size = New System.Drawing.Size(109, 24)
        Me.TranImportBtn.TabIndex = 0
        Me.TranImportBtn.Text = "Import"
        Me.TranImportBtn.UseVisualStyleBackColor = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.DividendTempltBtn)
        Me.GroupBox3.Controls.Add(Me.ImportDividendBtn)
        Me.GroupBox3.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(4, 123)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(447, 56)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Dividend"
        '
        'DividendTempltBtn
        '
        Me.DividendTempltBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.DividendTempltBtn.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DividendTempltBtn.Location = New System.Drawing.Point(316, 19)
        Me.DividendTempltBtn.Name = "DividendTempltBtn"
        Me.DividendTempltBtn.Size = New System.Drawing.Size(124, 24)
        Me.DividendTempltBtn.TabIndex = 1
        Me.DividendTempltBtn.Text = "Template"
        Me.DividendTempltBtn.UseVisualStyleBackColor = False
        '
        'ImportDividendBtn
        '
        Me.ImportDividendBtn.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.ImportDividendBtn.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ImportDividendBtn.Location = New System.Drawing.Point(136, 19)
        Me.ImportDividendBtn.Name = "ImportDividendBtn"
        Me.ImportDividendBtn.Size = New System.Drawing.Size(109, 24)
        Me.ImportDividendBtn.TabIndex = 0
        Me.ImportDividendBtn.Text = "Import"
        Me.ImportDividendBtn.UseVisualStyleBackColor = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(456, 220)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.StatusLbl)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "Excel Import Utility"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ImportBtn As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents StatusLbl As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TranImportBtn As System.Windows.Forms.Button
    Friend WithEvents TransTempltBtn As System.Windows.Forms.Button
    Friend WithEvents VoucherDateField As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents DividendTempltBtn As System.Windows.Forms.Button
    Friend WithEvents ImportDividendBtn As System.Windows.Forms.Button

End Class
