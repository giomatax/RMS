Imports System.Runtime.InteropServices
Imports System.Net.Mail

Imports System.Data.OleDb

Imports System.Reflection

Public Class Form1

    Inherits System.Windows.Forms.Form
    Dim WithEvents HidK As New Keyboard
    Private FRun As Boolean = False
    Dim s As String

    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer3 As System.Windows.Forms.Timer
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ბეჭდვაToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents TextP02 As System.Windows.Forms.TextBox
    Friend WithEvents TextP01 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents TabPage7 As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DateTimePickerხელშდამთავრება As System.Windows.Forms.DateTimePicker
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents TextAPEX01 As System.Windows.Forms.TextBox
    Friend WithEvents TextAPEX03 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TabControl2 As System.Windows.Forms.TabControl
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePickerხელშდაწყება As System.Windows.Forms.DateTimePicker
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents გადახდაToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
    Friend WithEvents SearchTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker5 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents ContextMenuStrip3 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents გადახდისთანხისკორექტირებაToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents გადახდისთარიღისკორექტირებაToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RMS As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenuforIcon As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents პროგრამისშესახებToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents გამოსვლაToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents sfdImage As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle19 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle20 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle21 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle17 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle18 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.sfdImage = New System.Windows.Forms.SaveFileDialog()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer3 = New System.Windows.Forms.Timer(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextP02 = New System.Windows.Forms.TextBox()
        Me.TextP01 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.TabPage7 = New System.Windows.Forms.TabPage()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.ContextMenuStrip3 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.DateTimePicker5 = New System.Windows.Forms.DateTimePicker()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.SearchTextBox = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextAPEX03 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextAPEX01 = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DateTimePickerხელშდამთავრება = New System.Windows.Forms.DateTimePicker()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.DateTimePickerხელშდაწყება = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.TabControl2 = New System.Windows.Forms.TabControl()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.გადახდაToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.გადახდისთანხისკორექტირებაToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.გადახდისთარიღისკორექტირებაToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.ბეჭდვაToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RMS = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenuforIcon = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.პროგრამისშესახებToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.გამოსვლაToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage7.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip3.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.TabControl2.SuspendLayout()
        Me.ContextMenuforIcon.SuspendLayout()
        Me.SuspendLayout()
        '
        'sfdImage
        '
        Me.sfdImage.FileName = "Webcam1"
        Me.sfdImage.Filter = "Bitmap|*.bmp"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 632)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1199, 22)
        Me.StatusStrip1.TabIndex = 21
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'Timer1
        '
        Me.Timer1.Interval = 300
        '
        'Timer3
        '
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Enabled = False
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ბეჭდვაToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(157, 26)
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "check-user-icon (1).png")
        Me.ImageList1.Images.SetKeyName(1, "Apps-preferences-web-browser-identification-icon.png")
        Me.ImageList1.Images.SetKeyName(2, "Ok-icon.png")
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 60000
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TableLayoutPanel2)
        Me.TabPage1.ImageIndex = 1
        Me.TabPage1.Location = New System.Drawing.Point(4, 31)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(1191, 597)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "შიდა კომპანიების დამატება"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.Panel3, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Panel4, 0, 1)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 2
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 180.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(1191, 597)
        Me.TableLayoutPanel2.TabIndex = 28
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Controls.Add(Me.TextP02)
        Me.Panel3.Controls.Add(Me.Button5)
        Me.Panel3.Controls.Add(Me.TextP01)
        Me.Panel3.Controls.Add(Me.Button4)
        Me.Panel3.Controls.Add(Me.Label2)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(3, 3)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1185, 174)
        Me.Panel3.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 12)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(166, 13)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "ჩვენი კომპანიის დასახელება"
        '
        'TextP02
        '
        Me.TextP02.Location = New System.Drawing.Point(8, 72)
        Me.TextP02.Name = "TextP02"
        Me.TextP02.Size = New System.Drawing.Size(286, 20)
        Me.TextP02.TabIndex = 20
        '
        'TextP01
        '
        Me.TextP01.Location = New System.Drawing.Point(8, 28)
        Me.TextP01.Name = "TextP01"
        Me.TextP01.Size = New System.Drawing.Size(286, 20)
        Me.TextP01.TabIndex = 19
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(196, 13)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "ჩვენი კომპანიის საიდენტიფიკაციო"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.DataGridView2)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(3, 183)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1185, 411)
        Me.Panel4.TabIndex = 1
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle11
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView2.DefaultCellStyle = DataGridViewCellStyle12
        Me.DataGridView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView2.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        DataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle19.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle19.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle19.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle19.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle19.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.RowHeadersDefaultCellStyle = DataGridViewCellStyle19
        Me.DataGridView2.Size = New System.Drawing.Size(1185, 411)
        Me.DataGridView2.TabIndex = 25
        '
        'TabPage7
        '
        Me.TabPage7.Controls.Add(Me.TableLayoutPanel1)
        Me.TabPage7.ImageIndex = 2
        Me.TabPage7.Location = New System.Drawing.Point(4, 31)
        Me.TabPage7.Name = "TabPage7"
        Me.TabPage7.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage7.Size = New System.Drawing.Size(1191, 597)
        Me.TabPage7.TabIndex = 1
        Me.TabPage7.Text = "ლიმიტის გაწერა"
        Me.TabPage7.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.DataGridView3, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel2, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 4
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 247.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 17.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1185, 591)
        Me.TableLayoutPanel1.TabIndex = 28
        '
        'DataGridView3
        '
        Me.DataGridView3.AllowUserToAddRows = False
        Me.DataGridView3.AllowUserToDeleteRows = False
        DataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle20.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle20.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle20.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle20.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle20.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView3.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle20
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.ContextMenuStrip = Me.ContextMenuStrip3
        DataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle21.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle21.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle21.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle21.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle21.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView3.DefaultCellStyle = DataGridViewCellStyle21
        Me.DataGridView3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView3.Location = New System.Drawing.Point(3, 413)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.ReadOnly = True
        Me.DataGridView3.Size = New System.Drawing.Size(1179, 157)
        Me.DataGridView3.TabIndex = 15
        '
        'ContextMenuStrip3
        '
        Me.ContextMenuStrip3.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.ContextMenuStrip3.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip3.Size = New System.Drawing.Size(181, 26)
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label21)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.Label17)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.ComboBox3)
        Me.Panel1.Controls.Add(Me.ComboBox2)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.Label10)
        Me.Panel1.Controls.Add(Me.DateTimePicker5)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.SearchTextBox)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.TextBox3)
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.ComboBox1)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.TextAPEX03)
        Me.Panel1.Controls.Add(Me.TextBox2)
        Me.Panel1.Controls.Add(Me.TextAPEX01)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.DateTimePickerხელშდამთავრება)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.DateTimePickerხელშდაწყება)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(3, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1179, 241)
        Me.Panel1.TabIndex = 0
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Location = New System.Drawing.Point(801, 96)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(59, 13)
        Me.Label21.TabIndex = 55
        Me.Label21.Text = "თეთრები"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(801, 62)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(61, 13)
        Me.Label20.TabIndex = 54
        Me.Label20.Text = "მწვანეები"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(801, 32)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(59, 13)
        Me.Label19.TabIndex = 53
        Me.Label19.Text = "წითლები"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ForeColor = System.Drawing.Color.Red
        Me.Label13.Location = New System.Drawing.Point(904, 35)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(45, 13)
        Me.Label13.TabIndex = 52
        Me.Label13.Text = "Label13"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label18.Location = New System.Drawing.Point(904, 88)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(45, 13)
        Me.Label18.TabIndex = 51
        Me.Label18.Text = "Label18"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.ForeColor = System.Drawing.Color.LawnGreen
        Me.Label17.Location = New System.Drawing.Point(904, 62)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(45, 13)
        Me.Label17.TabIndex = 50
        Me.Label17.Text = "Label17"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(14, 186)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(103, 13)
        Me.Label12.TabIndex = 48
        Me.Label12.Text = "მეიჯარე კომპანია"
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Location = New System.Drawing.Point(17, 202)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(194, 21)
        Me.ComboBox3.TabIndex = 47
        '
        'ComboBox2
        '
        Me.ComboBox2.AutoCompleteCustomSource.AddRange(New String() {"GEL", "USD", "EUR", "RUR", "GBP", "LIRA"})
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"GEL", "USD", "EUR", "RUR", "GPB"})
        Me.ComboBox2.Location = New System.Drawing.Point(446, 131)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(50, 21)
        Me.ComboBox2.TabIndex = 46
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(443, 115)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(53, 13)
        Me.Label11.TabIndex = 45
        Me.Label11.Text = "ვალუტა"
        '
        'Label10
        '
        Me.Label10.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label10.BackColor = System.Drawing.Color.DimGray
        Me.Label10.Location = New System.Drawing.Point(-47, 173)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(1837, 1)
        Me.Label10.TabIndex = 44
        '
        'DateTimePicker5
        '
        Me.DateTimePicker5.CustomFormat = "yyyy MMMM"
        Me.DateTimePicker5.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!)
        Me.DateTimePicker5.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker5.Location = New System.Drawing.Point(508, 204)
        Me.DateTimePicker5.Margin = New System.Windows.Forms.Padding(4)
        Me.DateTimePicker5.Name = "DateTimePicker5"
        Me.DateTimePicker5.ShowUpDown = True
        Me.DateTimePicker5.Size = New System.Drawing.Size(183, 19)
        Me.DateTimePicker5.TabIndex = 43
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!)
        Me.Label9.Location = New System.Drawing.Point(505, 186)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 13)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "პერიოდი"
        '
        'SearchTextBox
        '
        Me.SearchTextBox.Location = New System.Drawing.Point(233, 204)
        Me.SearchTextBox.Name = "SearchTextBox"
        Me.SearchTextBox.Size = New System.Drawing.Size(251, 20)
        Me.SearchTextBox.TabIndex = 40
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(230, 188)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(157, 13)
        Me.Label8.TabIndex = 41
        Me.Label8.Text = "გვარი, სახელი ან პირადი N"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(29, 16)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(103, 13)
        Me.Label14.TabIndex = 10
        Me.Label14.Text = "მეიჯარე კომპანია"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(559, 14)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(164, 13)
        Me.Label15.TabIndex = 8
        Me.Label15.Text = "მოიჯარეს საიდენტიფიკაციო"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(562, 78)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(183, 20)
        Me.TextBox3.TabIndex = 8
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(313, 14)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(134, 13)
        Me.Label16.TabIndex = 6
        Me.Label16.Text = "მოიჯარეს დასახელება"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(32, 32)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(251, 21)
        Me.ComboBox1.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(559, 62)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(186, 13)
        Me.Label7.TabIndex = 26
        Me.Label7.Text = "რამდენი დღით ადრე შეამოწმოს"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 59)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(201, 13)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "ხელშეკრულების დაწყების თარიღი"
        '
        'TextAPEX03
        '
        Me.TextAPEX03.Location = New System.Drawing.Point(316, 33)
        Me.TextAPEX03.Name = "TextAPEX03"
        Me.TextAPEX03.Size = New System.Drawing.Size(221, 20)
        Me.TextAPEX03.TabIndex = 2
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(316, 132)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(121, 20)
        Me.TextBox2.TabIndex = 7
        '
        'TextAPEX01
        '
        Me.TextAPEX01.Location = New System.Drawing.Point(562, 33)
        Me.TextAPEX01.Name = "TextAPEX01"
        Me.TextAPEX01.Size = New System.Drawing.Size(183, 20)
        Me.TextAPEX01.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(313, 115)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(124, 13)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "გადასახდელი თანხა"
        '
        'DateTimePickerხელშდამთავრება
        '
        Me.DateTimePickerხელშდამთავრება.CustomFormat = "dd MMMM yyyy"
        Me.DateTimePickerხელშდამთავრება.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerხელშდამთავრება.Location = New System.Drawing.Point(32, 136)
        Me.DateTimePickerხელშდამთავრება.Name = "DateTimePickerხელშდამთავრება"
        Me.DateTimePickerხელშდამთავრება.Size = New System.Drawing.Size(221, 20)
        Me.DateTimePickerხელშდამთავრება.TabIndex = 5
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(316, 76)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(121, 20)
        Me.TextBox1.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(313, 59)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(110, 13)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "გადახდის თარიღი"
        '
        'DateTimePickerხელშდაწყება
        '
        Me.DateTimePickerხელშდაწყება.CustomFormat = "dd MMMM yyyy"
        Me.DateTimePickerხელშდაწყება.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePickerხელშდაწყება.Location = New System.Drawing.Point(32, 75)
        Me.DateTimePickerხელშდაწყება.Name = "DateTimePickerხელშდაწყება"
        Me.DateTimePickerხელშდაწყება.Size = New System.Drawing.Size(221, 20)
        Me.DateTimePickerხელშდაწყება.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(29, 115)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(228, 13)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "ხელშეკრულების დამთავრების  თარიღი"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.DataGridView1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(3, 250)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1179, 157)
        Me.Panel2.TabIndex = 1
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        DataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle17.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle17.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle17.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle17.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle17.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle17
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.ContextMenuStrip = Me.ContextMenuStrip2
        DataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle18.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle18.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle18.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle18.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle18.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle18
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.Size = New System.Drawing.Size(1179, 157)
        Me.DataGridView1.TabIndex = 14
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.გადახდაToolStripMenuItem, Me.გადახდისთანხისკორექტირებაToolStripMenuItem, Me.გადახდისთარიღისკორექტირებაToolStripMenuItem})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(279, 70)
        '
        'TabControl2
        '
        Me.TabControl2.Controls.Add(Me.TabPage7)
        Me.TabControl2.Controls.Add(Me.TabPage1)
        Me.TabControl2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl2.ImageList = Me.ImageList1
        Me.TabControl2.Location = New System.Drawing.Point(0, 0)
        Me.TabControl2.Name = "TabControl2"
        Me.TabControl2.SelectedIndex = 0
        Me.TabControl2.Size = New System.Drawing.Size(1199, 632)
        Me.TabControl2.TabIndex = 33
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Image = Global.RMS.My.Resources.Resources.Delete
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(180, 22)
        Me.ToolStripMenuItem1.Text = "გადახდის წაშლა"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Image = Global.RMS.My.Resources.Resources.Add
        Me.Button2.Location = New System.Drawing.Point(562, 129)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(183, 23)
        Me.Button2.TabIndex = 9
        Me.Button2.Text = "დამატება"
        Me.Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = Global.RMS.My.Resources.Resources.Synchronize
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(713, 204)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(94, 22)
        Me.Button1.TabIndex = 16
        Me.Button1.Text = "განახლება"
        Me.Button1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button1.UseVisualStyleBackColor = True
        '
        'გადახდაToolStripMenuItem
        '
        Me.გადახდაToolStripMenuItem.Image = Global.RMS.My.Resources.Resources.Add
        Me.გადახდაToolStripMenuItem.Name = "გადახდაToolStripMenuItem"
        Me.გადახდაToolStripMenuItem.Size = New System.Drawing.Size(278, 22)
        Me.გადახდაToolStripMenuItem.Text = "გადახდა"
        '
        'გადახდისთანხისკორექტირებაToolStripMenuItem
        '
        Me.გადახდისთანხისკორექტირებაToolStripMenuItem.Image = Global.RMS.My.Resources.Resources.Synchronize
        Me.გადახდისთანხისკორექტირებაToolStripMenuItem.Name = "გადახდისთანხისკორექტირებაToolStripMenuItem"
        Me.გადახდისთანხისკორექტირებაToolStripMenuItem.Size = New System.Drawing.Size(278, 22)
        Me.გადახდისთანხისკორექტირებაToolStripMenuItem.Text = "გადახდის თანხის კორექტირება"
        '
        'გადახდისთარიღისკორექტირებაToolStripMenuItem
        '
        Me.გადახდისთარიღისკორექტირებაToolStripMenuItem.Image = Global.RMS.My.Resources.Resources.calendar_icon
        Me.გადახდისთარიღისკორექტირებაToolStripMenuItem.Name = "გადახდისთარიღისკორექტირებაToolStripMenuItem"
        Me.გადახდისთარიღისკორექტირებაToolStripMenuItem.Size = New System.Drawing.Size(278, 22)
        Me.გადახდისთარიღისკორექტირებაToolStripMenuItem.Text = "გადახდის თარიღის კორექტირება"
        '
        'Button5
        '
        Me.Button5.Image = Global.RMS.My.Resources.Resources.Add
        Me.Button5.Location = New System.Drawing.Point(202, 110)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(92, 23)
        Me.Button5.TabIndex = 23
        Me.Button5.Text = "დამატება"
        Me.Button5.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Image = Global.RMS.My.Resources.Resources.Synchronize
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(8, 110)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(93, 23)
        Me.Button4.TabIndex = 27
        Me.Button4.Text = "განახლება"
        Me.Button4.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button4.UseVisualStyleBackColor = True
        '
        'ბეჭდვაToolStripMenuItem
        '
        Me.ბეჭდვაToolStripMenuItem.Enabled = False
        Me.ბეჭდვაToolStripMenuItem.Image = Global.RMS.My.Resources.Resources.jumpicon
        Me.ბეჭდვაToolStripMenuItem.Name = "ბეჭდვაToolStripMenuItem"
        Me.ბეჭდვაToolStripMenuItem.Size = New System.Drawing.Size(156, 22)
        Me.ბეჭდვაToolStripMenuItem.Text = "გადმოვადება"
        '
        'RMS
        '
        Me.RMS.ContextMenuStrip = Me.ContextMenuforIcon
        Me.RMS.Icon = CType(resources.GetObject("RMS.Icon"), System.Drawing.Icon)
        Me.RMS.Text = "RMS"
        Me.RMS.Visible = True
        '
        'ContextMenuforIcon
        '
        Me.ContextMenuforIcon.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.პროგრამისშესახებToolStripMenuItem, Me.გამოსვლაToolStripMenuItem})
        Me.ContextMenuforIcon.Name = "ContextMenuforIcon"
        Me.ContextMenuforIcon.Size = New System.Drawing.Size(192, 70)
        '
        'პროგრამისშესახებToolStripMenuItem
        '
        Me.პროგრამისშესახებToolStripMenuItem.Name = "პროგრამისშესახებToolStripMenuItem"
        Me.პროგრამისშესახებToolStripMenuItem.Size = New System.Drawing.Size(191, 22)
        Me.პროგრამისშესახებToolStripMenuItem.Text = "პროგრამის შესახებ"
        '
        'გამოსვლაToolStripMenuItem
        '
        Me.გამოსვლაToolStripMenuItem.Name = "გამოსვლაToolStripMenuItem"
        Me.გამოსვლაToolStripMenuItem.Size = New System.Drawing.Size(191, 22)
        Me.გამოსვლაToolStripMenuItem.Text = "გამოსვლა"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1199, 654)
        Me.Controls.Add(Me.TabControl2)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Rental Management System"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage7.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip3.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.TabControl2.ResumeLayout(False)
        Me.ContextMenuforIcon.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    'Const cnnstr As String = "host=213.131.34.174;database=adem_prod5;User ID=adempiere;password=adempiere"
    Const WM_CAP As Short = &H400S

    Const WM_CAP_DRIVER_CONNECT As Integer = WM_CAP + 10
    Const WM_CAP_DRIVER_DISCONNECT As Integer = WM_CAP + 11
    Const WM_CAP_EDIT_COPY As Integer = WM_CAP + 30

    Const WM_CAP_SET_PREVIEW As Integer = WM_CAP + 50
    Const WM_CAP_SET_PREVIEWRATE As Integer = WM_CAP + 52
    Const WM_CAP_SET_SCALE As Integer = WM_CAP + 53
    Const WS_CHILD As Integer = &H40000000
    Const WS_VISIBLE As Integer = &H10000000
    Const SWP_NOMOVE As Short = &H2S
    Const SWP_NOSIZE As Short = 1
    Const SWP_NOZORDER As Short = &H4S
    Const HWND_BOTTOM As Short = 1

    Dim iDevice As Integer = 0 ' Current device ID
    Dim hHwnd As Integer ' Handle to preview window

    Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, _
        <MarshalAs(UnmanagedType.AsAny)> ByVal lParam As Object) As Integer

    Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Integer, _
        ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, _
        ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Declare Function DestroyWindow Lib "user32" (ByVal hndw As Integer) As Boolean

    Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
        (ByVal lpszWindowName As String, ByVal dwStyle As Integer, _
        ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, _
        ByVal nHeight As Short, ByVal hWndParent As Integer, _
        ByVal nID As Integer) As Integer

    Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Short, _
        ByVal lpszName As String, ByVal cbName As Integer, ByVal lpszVer As String, _
        ByVal cbVer As Integer) As Boolean

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'HookModule.UnhookKeyboard()
        End
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        HidK.DiposeHook()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        EnableDoubleBuffered(DataGridView1)
        EnableDoubleBuffered(DataGridView3)


        Me.Text = "Rental Management System -" & Assembly.GetExecutingAssembly().GetName().Version.ToString()
        FRun = True
        Dim inir As New Org.Mentalis.Files.IniReader(Application.StartupPath & "\BARCODE.ini")
        HidK.CreateHook()
        BarCode1 = ""
        KeyCount1 = 0
        StartChar = False
        CharLenght = 0

        TabNom = 4
        FormIsActive = True
        DateTimePickerხელშდამთავრება.Value = Today
        FillFuelCombo1(ComboBox1)


        FillFuelCombo2(ComboBox3)
        ' FillLimitGrig()
        FRun = False
        FillLimitGrig()
    End Sub
    Private Sub FillLimitGrig()
        Dim asa As Devart.Data.PostgreSql.PgSqlConnection
        asa = New Devart.Data.PostgreSql.PgSqlConnection
        Dim MQueryText As String
        asa.Unicode = True
        asa.Host = PSG_Host
        asa.Database = PSG_DBName
        asa.Port = PSG_Port
        asa.UserId = PSG_User
        asa.Password = PSG_Pass
        asa.Open()







        Dim stdate1 As String = Format(DateTimePickerხელშდამთავრება.Value, "yyyy-MM-dd")
        Dim stdate2 As String = Format(DateTimePickerხელშდამთავრება.Value.AddDays(1), "yyyy-MM-dd")










        If (SearchTextBox.Text.Length > 0) Then

            Debug.Print(SearchTextBox.Text.Length)


            If (ComboBox3.Text.Length > 0) Then
                Debug.Print(ComboBox3.Text.Length)
                MQueryText = "select i.dali_income_id,d.company_name,d.tax_id,icompany_name,i.itax_id,i.ipay_amount,i.currency,i.ipay_day,i.isactive,i.cstart_day,i.cend_day,i.alertday from dali_income  i left join dali_company d on i.dali_company_id=d.dali_company_id where     d.company_name='" & ComboBox3.Text & "' and (i.icompany_name like   '%" & SearchTextBox.Text & "%' or i.itax_id like '%" & SearchTextBox.Text & "%')"
                Debug.Print(MQueryText)
            Else
                MQueryText = "select i.dali_income_id,d.company_name,d.tax_id,icompany_name,i.itax_id,i.ipay_amount,i.currency,i.ipay_day,i.isactive,i.cstart_day,i.cend_day,i.alertday from dali_income  i left join dali_company d on i.dali_company_id=d.dali_company_id where i.icompany_name like  '%" & SearchTextBox.Text & "%' or i.itax_id like '%" & SearchTextBox.Text & "%' "

            End If

        Else



            If (ComboBox3.Text.Length > 0) Then
                MQueryText = "select i.dali_income_id,d.company_name,d.tax_id,icompany_name,i.itax_id,i.ipay_amount,i.currency,i.ipay_day,i.isactive,i.cstart_day,i.cend_day,i.alertday from dali_income  i left join dali_company d on i.dali_company_id=d.dali_company_id where  d.company_name='" & ComboBox3.Text & "' "

            Else

                MQueryText = "select i.dali_income_id,d.company_name,d.tax_id,icompany_name,i.itax_id,i.ipay_amount,i.currency,i.ipay_day,i.isactive,i.cstart_day,i.cend_day,i.alertday from dali_income  i left join dali_company d on i.dali_company_id=d.dali_company_id "





            End If
        End If



        Dim asa3 As New Devart.Data.PostgreSql.PgSqlDataAdapter(MQueryText, asa)
        Dim ds As New DataSet
        asa3.Fill(ds, "dali_income")
        asa.Close()

        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "dali_income"
        DataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        ' DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(0).Visible = False


        DataGridView1.Columns(1).HeaderText = "მეიჯარე კომპანია"
        DataGridView1.Columns(1).Width = 120

        DataGridView1.Columns(2).HeaderText = "მეიჯარე საიდენტიფიკაციო"
        DataGridView1.Columns(2).Width = 120

        DataGridView1.Columns(3).HeaderText = "მოიჯარე კომპანია"
        DataGridView1.Columns(3).Width = 120


        DataGridView1.Columns(4).HeaderText = "მოიჯარე საიდენტიფიკაციო"
        DataGridView1.Columns(4).Width = 120


        DataGridView1.Columns(5).HeaderText = "გადასახდელი თანხა"
        DataGridView1.Columns(5).Width = 120




        DataGridView1.Columns(6).HeaderText = "ვალუტა"
        DataGridView1.Columns(6).Width = 60



        DataGridView1.Columns(7).HeaderText = "გადახდის რიცხვი"
        DataGridView1.Columns(7).Width = 120

        DataGridView1.Columns(8).Visible = False

        'DataGridView1.Columns(8).HeaderText = "ამ თვეში გადახდილი თანხა"
        'DataGridView1.Columns(8).Width = 120




        DataGridView1.Columns(9).HeaderText = "ხელშეკრულების დაწყების თარიღი"
        DataGridView1.Columns(9).Width = 120

        DataGridView1.Columns(10).HeaderText = "ხელშეკრულების დამთავრების თარიღი"
        DataGridView1.Columns(10).Width = 120


        DataGridView1.Columns(11).HeaderText = "რამდენი დღით ადრე შეამოწმოს"
        DataGridView1.Columns(11).Width = 120



        DataGridView1.RowHeadersWidth = 5




        TaskStatusRefresh()



    End Sub
    Private Sub TaskStatusRefresh()


        Dim Rnum As Integer = 0
        Dim Gnum As Integer = 0
        Dim Wnum As Integer = 0

        If DataGridView1.RowCount > 0 Then

            Dim PeriodStr2 As String = DateTimePicker5.Value.Year.ToString & IIf(Len(DateTimePicker5.Value.Month.ToString) = 1, "0" & DateTimePicker5.Value.Month.ToString, DateTimePicker5.Value.Month.ToString)




            'select *  from dali_income i   left join dali_income_line_group_v dl on dl.dali_income_id=i.dali_income_id  and dl.period=202207  where   coalesce(dl.sumamt,0)<i.ipay_amount

            Dim raga2 As DataSet = FillPSQLDataSet2("select *  from dali_income i   left join dali_income_line_group_v dl on dl.dali_income_id=i.dali_income_id  and dl.period=" & PeriodStr2 & "  where   coalesce(dl.sumamt,0)<i.ipay_amount ")
            Dim PR2Rows() As System.Data.DataRow = Nothing

            Dim raga As DataSet = FillPSQLDataSet2("select i.dali_income_id  from dali_income i left join dali_income_line_group_v dl on dl.dali_income_id=i.dali_income_id  and dl.period=" & PeriodStr2 & " where  CAST((to_char(date_trunc('day'::text, now()),'dd') ) AS DECIMAL)>(i.ipay_day-i.alertday) and coalesce(dl.sumamt,0)<i.ipay_amount")
            Dim PR1Rows() As System.Data.DataRow = Nothing

            For ox As Integer = 0 To DataGridView1.RowCount - 1



                PR2Rows = raga2.Tables(0).Select("dali_income_id=" & DataGridView1.Item(0, ox).Value)
                If PR2Rows.Length = 1 Then
                    DataGridView1.Rows(ox).DefaultCellStyle.BackColor = Color.White
                    Wnum = Wnum + 1
                Else
                    DataGridView1.Rows(ox).DefaultCellStyle.BackColor = Color.LightGreen
                    Gnum = Gnum + 1

                End If



                PR1Rows = raga.Tables(0).Select("dali_income_id=" & DataGridView1.Item(0, ox).Value)
                If PR1Rows.Length = 1 Then



                    DataGridView1.Rows(ox).DefaultCellStyle.BackColor = Color.LightPink

                    Rnum = Rnum + 1








                End If






            Next
        End If


        Label13.Text = Rnum
        Label17.Text = Gnum
        Label18.Text = Wnum


    End Sub











    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer3.Enabled = True
        Timer1.Enabled = False
    End Sub

    Private Sub Timer3_Tick(sender As System.Object, e As System.EventArgs) Handles Timer3.Tick
        BarCode1 = ""
        KeyCount1 = 0
        BarCode2 = ""
        KeyCount2 = 0
        BarCode3 = ""
        KeyCount3 = 0

        Timer3.Enabled = False
    End Sub

    Private Sub TabPage6_Enter(sender As Object, e As System.EventArgs)
        TabNom = 4
    End Sub
    Private Sub TabPage7_Enter(sender As Object, e As System.EventArgs)
        TabNom = 5


    End Sub



    Private Sub SendMailReport(ByVal MailText As String)
        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Credentials = New Net.NetworkCredential("sakcable.adempiere@gmail.com", "Fmazag4e")
            Smtp_Server.Host = "smtp.gmail.com"
            e_mail = New MailMessage()
            e_mail.From = New MailAddress("sakcable.adempiere@gmail.com", "SAKCABLE Report")
            e_mail.To.Add("k.jananashvili@sakcable.ge")
            'e_mail.To.Add("m.sharashidze@sakcable.ge")
            'e_mail.To.Add("2easy4roland@gmail.com")
            e_mail.To.Add("g.matakheria@sakcable.ge")
            'e_mail.To.Add("ts.giorgi@gmail.com")
            e_mail.Subject = "ლიმიტის გაწერა - " & Format(Now(), "dd MMMM yyyy")
            e_mail.IsBodyHtml = True
            e_mail.Body = MailText
            Smtp_Server.Send(e_mail)
        Catch error_t As Exception
            Debug.Print(error_t.ToString)
            MsgBox(error_t.ToString)
        End Try
    End Sub
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs)
        Dim CntNameFA As String
        Dim CntIDFA As String
        CntNameFA = ""
        CntIDFA = ""
        ' If CheckShopUser(ApexUser) Then
        'If Val(TextAPEX02.Text) > OutAmt(ApexUser) Then
        '    MsgBox("თქვენ არ შეგიძლიათ " & OutAmt(ApexUser) & " ლარზე მეტი თანხის გაწერა", MsgBoxStyle.Exclamation)
        'Else
        CntIDFA = CheckCntID(TextAPEX01.Text, CntNameFA)
        If CntNameFA = "NULL" Then
            MsgBox("ხელშეკრულების ნომერი ვერ მოიძებნა!", MsgBoxStyle.Exclamation)
        Else
            If MsgBox("დარწმუნებული ხარ რომ გინდა გაუზარდო ლიმიტი " & CntNameFA & "-ს  თანხით?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If InsertInLimiter(CntIDFA, CntNameFA, "1", "1", TextAPEX03.Text) Then
                    ' SendMailReport(ApexUser & "-მა გაწერა " & TextAPEX02.Text & " ლარიანი ლიმიტი " & CntNameFA & "-ზე [ხელშეკრულება - " & CntIDFA & "] - (" & TextAPEX03.Text & ")")
                    MsgBox("ლიმიტი მოთხოვნილია", MsgBoxStyle.Information)
                Else
                    MsgBox("ლიმიტის მოთხოვნა ვერ მოხერხდა!", MsgBoxStyle.Exclamation)
                End If
            End If
        End If
        FillLimitGrig()
        'End If
        'Else
        'If CheckCntDate(Today) Then
        '    If Val(TextAPEX02.Text) > OutAmt(ApexUser) Then
        '        MsgBox("თქვენ არ შეგიძლიათ " & OutAmt(ApexUser) & " ლარზე მეტი თანხის გაწერა", MsgBoxStyle.Exclamation)
        '    Else
        '        CntIDFA = CheckCntID(TextAPEX01.Text, CntNameFA)
        '        If CntNameFA = "NULL" Then
        '            MsgBox("ხელშეკრულების ნომერი ვერ მოიძებნა!", MsgBoxStyle.Exclamation)
        '        Else
        '            If MsgBox("დარწმუნებული ხარ რომ გინდა გაუზარდო ლიმიტი " & CntNameFA & "-ს " & TextAPEX02.Text & " თანხით?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
        '                If UpdateCntAMT(CntIDFA, TextAPEX02.Text, TextAPEX03.Text, ApexUser) Then
        '                    SendMailReport(ApexUser & "-მა გაწერა " & TextAPEX02.Text & " ლარიანი ლიმიტი " & CntNameFA & "-ზე [ხელშეკრულება - " & CntIDFA & "] - (" & TextAPEX03.Text & ")")
        '                    MsgBox("ლიმიტი გაწერილია", MsgBoxStyle.Information)
        '                Else
        '                    MsgBox("ლიმიტი ვარ გაიწერა!", MsgBoxStyle.Exclamation)
        '                End If
        '            End If
        '        End If
        '    End If
        'Else
        '    MsgBox("დღეს ლიმიტის გაწერა არ შეგიძლიათ!", MsgBoxStyle.Exclamation)
        'End If
        'End If



    End Sub

    'EmplID = CheckEmplID(cardcode)
    'If EmplID = "" Then
    '    ShowInfo(Format(TimeOfDay, "HH:mm:ss") & " - " & "ბარათის მფლობელი ვერ მოიძებნა")
    '    AddLog(GetCardCode & " - ბარათის მფლობელი ვერ მოიძებნა (ტაბელი)")
    '    GoTo bolo
    'End If
    'EmplName = GetEmplName(EmplID)
    'TextBox9.Text = EmplName
    'SaveEmployeeHistory(EmplID)
    'TextBox8.Text = Format(TimeOfDay, "HH:mm:ss")
    'TextBox7.Text = Format(Today, "dd.MM.yyyy")




    Private Sub InsertInTime1(ByVal BP_ID As String)
        'Dim a As Integer
        Dim commandstr As String
        Dim StDate1 As String

        Dim asa As Devart.Data.PostgreSql.PgSqlConnection
        Dim asa2 As Devart.Data.PostgreSql.PgSqlCommand
        asa = New Devart.Data.PostgreSql.PgSqlConnection
        asa2 = New Devart.Data.PostgreSql.PgSqlCommand
        asa.Host = PSG_Host
        asa.Database = PSG_DBName
        asa.Port = PSG_Port
        asa.UserId = PSG_User
        asa.Password = PSG_Pass
        asa.Unicode = True
        asa.Open()
        asa2.Connection() = asa
        StDate1 = Format(Today, "yyyy-MM-dd") & " " & Format(Now, " HH:mm:ss")
        commandstr = "INSERT INTO _saq_employee_timecard(_saq_employee_timecard_id, ad_client_id, ad_org_id, created, createdby, updated, updatedby, description, c_bpartner_id, intime, outtime, duration)" & _
          "VALUES (" & GetLastID() & ", 1000000, 0, '" & StDate1 & "', 100,'" & StDate1 & "', 100,'მოსვლა ბარათით T1', " & BP_ID & ", '" & StDate1 & "', null, null)"

        asa2.CommandText = commandstr
        Try
            If asa2.ExecuteNonQuery().ToString = "1" Then

            Else
                MsgBox("შეცდომა, ტაბელი ვერ დაიმახსოვრა")
            End If
        Catch myException As Devart.Data.PostgreSql.PgSqlException
            MsgBox(myException.Message)
        End Try

        asa.Close()
    End Sub
    Private Sub InsertOutTime(ByVal Rec_ID As String)
        'Dim a As Integer
        Dim commandstr As String
        Dim StDate1 As String

        Dim asa As Devart.Data.PostgreSql.PgSqlConnection
        Dim asa2 As Devart.Data.PostgreSql.PgSqlCommand
        asa = New Devart.Data.PostgreSql.PgSqlConnection
        asa2 = New Devart.Data.PostgreSql.PgSqlCommand
        asa.Host = PSG_Host
        asa.Database = PSG_DBName
        asa.Port = PSG_Port
        asa.UserId = PSG_User
        asa.Password = PSG_Pass
        asa.Unicode = True
        asa.Open()
        asa2.Connection() = asa
        StDate1 = Format(Today, "yyyy-MM-dd") & " " & Format(Now, " HH:mm:ss")
        commandstr = "UPDATE _saq_employee_timecard SET outtime='" & StDate1 & "' WHERE _saq_employee_timecard_id=" & Rec_ID
        '"UPDATE _saq_employee_timecard SET outtime='" & StDate1 & "' WHERE _saq_employee_timecard_id=" & Rec_ID
        asa2.CommandText = commandstr
        Try
            If asa2.ExecuteNonQuery().ToString = "1" Then

            Else
                MsgBox("შეცდომა, ტაბელი ვერ დაიმახსოვრა")
            End If
        Catch myException As Devart.Data.PostgreSql.PgSqlException
            MsgBox(myException.Message)
        End Try

        asa.Close()
    End Sub
    Private Function InsertDuration(ByVal Rec_ID As String) As String
        'Dim a As Integer
        Dim commandstr As String
        Dim StDate1 As String

        Dim asa As Devart.Data.PostgreSql.PgSqlConnection
        Dim asa2 As Devart.Data.PostgreSql.PgSqlCommand
        asa = New Devart.Data.PostgreSql.PgSqlConnection
        asa2 = New Devart.Data.PostgreSql.PgSqlCommand
        asa.Host = PSG_Host
        asa.Database = PSG_DBName
        asa.Port = PSG_Port
        asa.UserId = PSG_User
        asa.Password = PSG_Pass
        asa.Unicode = True
        asa.Open()
        asa2.Connection() = asa
        InsertDuration = ""
        StDate1 = Format(Today, "yyyy-MM-dd") & " " & Format(Now, " HH:mm:ss")
        commandstr = "UPDATE _saq_employee_timecard SET duration=outtime-intime WHERE _saq_employee_timecard_id=" & Rec_ID
        '"UPDATE _saq_employee_timecard SET outtime='" & StDate1 & "' WHERE _saq_employee_timecard_id=" & Rec_ID
        asa2.CommandText = commandstr
        Try
            If asa2.ExecuteNonQuery().ToString = "1" Then

            Else
                MsgBox("შეცდომა, ტაბელი ვერ დაიმახსოვრა")
            End If
        Catch myException As Devart.Data.PostgreSql.PgSqlException
            MsgBox(myException.Message)
        End Try
        InsertDuration = GetDuration(Rec_ID)
        asa.Close()
    End Function



    Public Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        FillLimitGrig()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex < 0 Then Exit Sub
        FillSpecLine1(DataGridView1.Item(0, e.RowIndex).Value)


    End Sub







    Private Sub FillSpecLine1(ByVal FSpecID As String)


        '   Debug.Print(FSpecID)

        Dim rowD As DataGridViewRow = DataGridView2.RowTemplate
        rowD.Height = 35

        Dim SLDsaet As DataSet = FillPSQLDataSet("SELECT * FROM dali_income_line Where dali_income_id =" & FSpecID & " Order By ipay_day")

        Debug.Print(FSpecID)
        DataGridView3.DataSource = SLDsaet.Tables(0)


        '  Debug.Print("aq ukve unda shevsebuliyo")

        DataGridView3.RowHeadersVisible = False




        DataGridView3.Columns(0).Visible = False
        DataGridView3.Columns(1).Visible = False
        DataGridView3.Columns(2).Visible = False


        DataGridView3.Columns(3).HeaderText = "პერიოდი"
        DataGridView3.Columns(3).Width = 120
        DataGridView3.Columns(4).Visible = False


        DataGridView3.Columns(5).HeaderText = "გადახდილი თანხა"
        DataGridView3.Columns(5).Width = 120
        DataGridView3.Columns(6).HeaderText = "გადახდის თარიღი"
        DataGridView3.Columns(6).Width = 120



    End Sub



    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub




    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        If TextP01.Text = "" Or TextP02.Text = "" Then
            MsgBox("ყველა ველი არ არის შევსილი!", MsgBoxStyle.Exclamation)
            Exit Sub
        Else
            Dim StDate1 As String = Format(Today, "yyyy-MM-dd") ' & " 00:00:00"
            InsertInTable("INSERT INTO dali_company(company_name,tax_id)" & _
                                  "VALUES ('" & TextP01.Text & "', '" & TextP02.Text & "')")
            FillPassGrig()
            'INSERT INTO tmap_pass(tmap_pass_type_id, create_user_id, approve_user_id, company_name, car_number, car_model, pass_status, in_status, pass_date) VALUES (?, ?, ?, ?, )
        End If
        TextP02.Text = ""
        TextP01.Text = ""
        FillFuelCombo1(ComboBox1)
    End Sub
    Private Sub FillPassGrig()
        Dim asa As Devart.Data.PostgreSql.PgSqlConnection
        asa = New Devart.Data.PostgreSql.PgSqlConnection
        Dim MQueryText As String
        asa.Unicode = True
        asa.Host = PSG_Host
        asa.Database = PSG_DBName
        asa.Port = PSG_Port
        asa.UserId = PSG_User
        asa.Password = PSG_Pass
        asa.Open()
        'Dim stdate1 As String = Format(DateTimePicker1.Value, "yyyy-MM-dd")
        'Dim stdate2 As String = Format(DateTimePicker1.Value.AddDays(1), "yyyy-MM-dd")

        MQueryText = "select  * from dali_company  "
        Dim asa3 As New Devart.Data.PostgreSql.PgSqlDataAdapter(MQueryText, asa)
        Dim ds As New DataSet
        asa3.Fill(ds, "dali_company")
        asa.Close()
        DataGridView2.DataSource = ds
        DataGridView2.DataMember = "dali_company"
        DataGridView2.RowHeadersVisible = False
        DataGridView2.Columns(0).Visible = False
        DataGridView2.Columns(1).Visible = False
        DataGridView2.Columns(2).HeaderText = "კომპანიის დასახელება"
        DataGridView2.Columns(2).Width = 250
        DataGridView2.Columns(3).HeaderText = "კომპანიის საიდენტიფიკაციო"
        DataGridView2.Columns(3).Width = 150

        'StockControlState()
    End Sub
    Private Sub StockControlState()
        Dim a As Integer

        For a = 0 To DataGridView2.RowCount - 1
            Select Case DataGridView2.Item(15, a).Value.ToString
                Case "N"
                    DataGridView2.Rows(a).DefaultCellStyle.BackColor = Color.White
                Case "P"
                    DataGridView2.Rows(a).DefaultCellStyle.BackColor = Color.Khaki
                Case "A"
                    DataGridView2.Rows(a).DefaultCellStyle.BackColor = Color.LightGreen
                Case "C"
                    DataGridView2.Rows(a).DefaultCellStyle.BackColor = Color.LightGray
            End Select
            If DataGridView2.Rows(a).Cells(11).Value.ToString = "გარეთ" Then
                DataGridView2.Rows(a).Cells(11).Style.BackColor = Color.Red
            Else
                DataGridView2.Rows(a).Cells(11).Style.BackColor = Color.Green
            End If
        Next

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        FillPassGrig()
    End Sub

    Private Sub ბეჭდვაToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ბეჭდვაToolStripMenuItem.Click
        If DataGridView2.RowCount = 0 Then Exit Sub
        Dim StDate1 As String = Format(Today, "yyyy-MM-dd") ' & " 00:00:00"
        If DataGridView2.Rows(DataGridView2.CurrentRow.Index).Cells(15).Value.ToString = "N" Then
            InsertInTable("UPDATE tmap_pass SET  pass_date='" & StDate1 & "',pass_no = " & GenPassNo() & " WHERE tmap_pass_id=" & DataGridView2.Item(0, DataGridView2.CurrentRow.Index).Value)
        Else
            MsgBox("საშვის სტატუსი უნდა იყოს ახალი!", MsgBoxStyle.Exclamation)
        End If
        FillPassGrig()
    End Sub

    Private Sub TabPage1_Enter(sender As Object, e As System.EventArgs) Handles TabPage1.Enter
        FillPassGrig()
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextAPEX01.Text = "" Or TextAPEX03.Text = "" Or TextBox1.Text = "" Or TextBox3.Text = "" Or TextBox2.Text = "" Or ComboBox2.Text = "" Then
            MsgBox("ყველა ველი არ არის შევსილი!", MsgBoxStyle.Exclamation)
            Exit Sub
        Else








            Debug.Print("შემოვიდა")

            Dim Get_DCOMPANY_ID As String = Get_DCOMPANY_ID_From_name(ComboBox1.Text)
            Debug.Print(Get_DCOMPANY_ID)

            Dim StDate1 As String = Format(Today, "yyyy-MM-dd") & " " & Format(Now, " HH:mm:ss")

            Debug.Print(StDate1)



            Dim Get_FQty As String = TextBox1.Text
            Debug.Print(Get_FQty)

            Dim Get_FOdometr As String = TextBox2.Text


            Debug.Print(Get_FOdometr)

            Dim DTimePickerდაწყება As String = Format(DateTimePickerხელშდაწყება.Value, "yyyy-MM-dd")
            Dim DTimePickerდამთავრება As String = Format(DateTimePickerხელშდამთავრება.Value, "yyyy-MM-dd")

            Debug.Print(DTimePickerდაწყება)

            Debug.Print(DTimePickerდამთავრება)


            InsertInTable("INSERT INTO dali_income( dali_company_id,icompany_name,itax_id,ipay_day,ipay_amount,cstart_day,cend_day,isactive,alertday,currency)  VALUES ('" & Get_DCOMPANY_ID & "', '" & TextAPEX03.Text & "','" & TextAPEX01.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & DTimePickerდაწყება & "','" & DTimePickerდამთავრება & "','Y','" & TextBox3.Text & "','" & ComboBox2.Text & "' );")


            AddLog("INSERT INTO dali_income( dali_company_id,icompany_name,itax_id,ipay_day,ipay_amount,cstart_day,cend_day,isactive,alertday)  VALUES ('" & Get_DCOMPANY_ID & "', '" & TextAPEX03.Text & "','" & TextAPEX01.Text & "','" & TextBox1.Text & "','" & TextBox2.Text & "','" & DTimePickerდაწყება & "','" & DTimePickerდამთავრება & "','Y','" & TextBox3.Text & "','" & ComboBox2.Text & "' );")


            Debug.Print("mimdinareobs inserti")

            MsgBox("წარმატებულად დაემატა კომპანია")
            AddLog("წარმატებულად დაემატა კომპანია")
            FillLimitGrig()
            Debug.Print("FillLimitGrig")

            TextAPEX03.Text = ""
            TextAPEX01.Text = ""
            TextBox1.Text = ""
            TextBox3.Text = ""
            TextBox2.Text = ""
            FillLimitGrig()

        End If
    End Sub

    Private Sub გადახდაToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles გადახდაToolStripMenuItem.Click

        Dali_income_id_selected = DataGridView1.Item(0, DataGridView1.SelectedCells.Item(0).RowIndex).Value

        Dali_income_ipay_amount_selected = DataGridView1.Item(5, DataGridView1.SelectedCells.Item(0).RowIndex).Value

        PAY.Show()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs)

    End Sub

    Private Sub SearchTextBox_TextChanged(sender As Object, e As EventArgs) Handles SearchTextBox.TextChanged
        FillLimitGrig()
    End Sub

    Private Sub MenuStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs)

    End Sub


    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click


 

        Dim ask As MsgBoxResult = MsgBox("წავშალოთ აღნიშნული ჩანაწერი?", MsgBoxStyle.YesNo)
        If ask = MsgBoxResult.Yes Then
            Dali_income_pay_line_selected = DataGridView3.Item(0, DataGridView3.SelectedCells.Item(0).RowIndex).Value

            Debug.Print(Dali_income_pay_line_selected)
            Dim inserttext As String

            inserttext = "DELETE FROM dali_income_line WHERE dali_income_line_id='" & Dali_income_pay_line_selected & "';"
            Debug.Print(inserttext)
            InsertInTable(inserttext, False)
        End If


       


    End Sub





    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

        If e.RowIndex < 0 Then Exit Sub
        Debug.Print(DataGridView1.Item(0, e.RowIndex).Value)

    End Sub

    Private Sub გადახდისთანხისკორექტირებაToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles გადახდისთანხისკორექტირებაToolStripMenuItem.Click
        Dim myNum As Integer
        Dim dali_id As Integer = DataGridView1.Item(0, DataGridView1.SelectedCells.Item(0).RowIndex).Value
        Dim inserttext As String

        Debug.Print(dali_id)
        AddLog(dali_id)
        Try

            myNum = InputBox("გთხოვთ შეიყვანოთ შესაცვლელი თანხა.")


            Debug.Print(myNum)

            AddLog(myNum)


            inserttext = "UPDATE dali_income   SET  ipay_amount='" & myNum & "'   WHERE dali_income_id='" & dali_id & "';"
            Debug.Print(inserttext)


            AddLog(inserttext)
            InsertInTable(inserttext, False)


        Catch

            MsgBox("ციფრის ფორმატი არასწორია.")

        End Try

        FillLimitGrig()
    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening

    End Sub

    Private Sub გადახდისთარიღისკორექტირებაToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles გადახდისთარიღისკორექტირებაToolStripMenuItem.Click
        Dim myNum As Integer
        Dim dali_id As Integer = DataGridView1.Item(0, DataGridView1.SelectedCells.Item(0).RowIndex).Value
        Dim inserttext As String

        Debug.Print(dali_id)
        AddLog(dali_id)
        Try

            myNum = InputBox("გთხოვთ გადახდის თარიღი")


            Debug.Print(myNum)
            AddLog(myNum)

            inserttext = "UPDATE dali_income   SET  ipay_day='" & myNum & "'   WHERE dali_income_id='" & dali_id & "';"
            Debug.Print(inserttext)
            AddLog(inserttext)
            InsertInTable(inserttext, False)


        Catch

            MsgBox("ციფრის ფორმატი არასწორია.")

        End Try

        FillLimitGrig()
    End Sub

    Private Sub ContextMenuforIcon_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuforIcon.Opening

    End Sub

    Private Sub პროგრამისშესახებToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles პროგრამისშესახებToolStripMenuItem.Click
        About.Show()
    End Sub
End Class
