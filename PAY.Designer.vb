<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PAY
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
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.DateTimePickerPD = New System.Windows.Forms.DateTimePicker()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.DateTimePicker5 = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!)
        Me.TextBox2.Location = New System.Drawing.Point(30, 164)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(166, 24)
        Me.TextBox2.TabIndex = 28
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!)
        Me.Label6.Location = New System.Drawing.Point(27, 138)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(169, 18)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "გადასახდელი თანხა"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!)
        Me.Label5.Location = New System.Drawing.Point(27, 74)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(149, 18)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "გადახდის თარიღი"
        '
        'DateTimePickerPD
        '
        Me.DateTimePickerPD.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 11.25!)
        Me.DateTimePickerPD.CustomFormat = "dd MMMM yyyy"
        Me.DateTimePickerPD.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!)
        Me.DateTimePickerPD.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerPD.Location = New System.Drawing.Point(30, 100)
        Me.DateTimePickerPD.Name = "DateTimePickerPD"
        Me.DateTimePickerPD.Size = New System.Drawing.Size(166, 24)
        Me.DateTimePickerPD.TabIndex = 29
        '
        'Button2
        '
        Me.Button2.Image = Global.RMS.My.Resources.Resources.Add
        Me.Button2.Location = New System.Drawing.Point(30, 208)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(166, 23)
        Me.Button2.TabIndex = 30
        Me.Button2.Text = "გადახდა"
        Me.Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DateTimePicker5
        '
        Me.DateTimePicker5.CustomFormat = "yyyy MMMM"
        Me.DateTimePicker5.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.DateTimePicker5.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker5.Location = New System.Drawing.Point(30, 33)
        Me.DateTimePicker5.Name = "DateTimePicker5"
        Me.DateTimePicker5.ShowUpDown = True
        Me.DateTimePicker5.Size = New System.Drawing.Size(166, 24)
        Me.DateTimePicker5.TabIndex = 32
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(27, 9)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 18)
        Me.Label3.TabIndex = 31
        Me.Label3.Text = "პერიოდი"
        '
        'PAY
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(225, 260)
        Me.Controls.Add(Me.DateTimePicker5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DateTimePickerPD)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "PAY"
        Me.Text = "PAY"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerPD As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents DateTimePicker5 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
