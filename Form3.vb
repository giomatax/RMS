Imports System
Imports System.IO
Imports System.Text
Imports System.Reflection

Public Class Form3


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '  OkClick()

        Dim ADpass As String
        Dim statustext As String
        ADpass = GetPass(TextBox1.Text, TmapUserID, TmapUserName, TmapUserRole, TmapUserBPartnerID)
        If ADpass = "NULL" Then
            MsgBox("მომხმარებელი არა სწორია!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        If TextBox2.Text = ADpass Then
            ADorgName = GetOrgname(TmapUserBPartnerID, ADorgID)
            TmapUserNameEN = TextBox1.Text

            ADUserID = GetUserID(TextBox1.Text)
            statustext = TextBox1.Text & " / " & ADorgName
            SaveSetting("BP Card", "Username", "Username", TextBox1.Text)
            Form1.Show()

            FormIsActive = True
            Me.Hide()





        Else
            MsgBox("პაროლი არა სწორია!", MsgBoxStyle.Exclamation)
        End If


    End Sub
    'Private Sub OkClick()
    '    Dim ADpass As String
    '    Dim statustext As String
    '    ADpass = GetPass(TextBox1.Text)
    '    If ADpass = "NULL" Then
    '        MsgBox("მომხმარებელი არა სწორია!", MsgBoxStyle.Exclamation)
    '        Exit Sub
    '    End If
    '    If TextBox2.Text = ADpass Then
    '        ADorgID = GetOrgID(TextBox1.Text)
    '        If ADorgID = 0 Then
    '            ADorgID = 1000000

    '            'Form2.Button2.Visible = True
    '        End If

    '        Select Case ADorgName
    '            Case "მაღაზია 1"
    '                RSUser = "saqmag1"
    '                RSPass = "saquser"
    '                ApexUser = "Shop1"
    '            Case "მაღაზია 2"
    '                RSUser = "saqmag2"
    '                RSPass = "saquser"
    '                ApexUser = "Shop2"
    '            Case "მაღაზია 3"
    '                RSUser = "saqmag3"
    '                RSPass = "saquser"
    '                ApexUser = "Shop3"
    '            Case "მაღაზია 4"
    '                RSUser = "saqmag4"
    '                RSPass = "saquser"
    '                ApexUser = "Shop4"
    '            Case "მაღაზია 5"
    '                RSUser = "saqmag5"
    '                RSPass = "saquser"
    '                ApexUser = "Shop5"
    '            Case "მაღაზია 6"
    '                RSUser = "saqmag6"
    '                RSPass = "saquser"
    '                ApexUser = "Shop6"
    '            Case "მაღაზია 8"
    '                RSUser = "saqmag8"
    '                RSPass = "saquser"
    '                ApexUser = "Shop8"
    '            Case "მაღაზია 9"
    '                RSUser = "saqmag9"
    '                RSPass = "saquser"
    '                ApexUser = "Shop9"
    '            Case "თბილისის საწყობი"
    '                RSUser = "saqmag77"
    '                RSPass = "saquser"
    '                ApexUser = "Shop7"
    '            Case "სს " & Chr(34) & "საქკაბელი" & Chr(34)
    '                RSUser = "ირაკლი"
    '                RSPass = "789"
    '        End Select
    '        ADUserID = GetUserID(TextBox1.Text)
    '        statustext = TextBox1.Text & " / " & ADorgName
    '        SaveSetting("BP Card", "Username", "Username", TextBox1.Text)
    '        Form1.Show()
    '        Form1.ToolStripStatusLabel1.Text = statustext
    '        FormIsActive = True
    '        Me.Hide()


    '    Else
    '        MsgBox("პაროლი არა სწორია!", MsgBoxStyle.Exclamation)
    '    End If
    'End Sub
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.CenterToScreen()
        TextBox1.TextAlign = HorizontalAlignment.Center
        TextBox2.TextAlign = HorizontalAlignment.Center

        Label5.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString()
        FormIsActive = False
        Me.AcceptButton = Button1
        TextBox1.Text = GetSetting("BP Card", "Username", "Username")
        Dim inir As New Org.Mentalis.Files.IniReader(Application.StartupPath & "\BARCODE.ini")
        'taskBar = FindWindow("Shell_traywnd", "")


        IniSet = inir.ReadString("Section1", "SET")
        'If IniSet = "1" Then
        '    Form1.Width = 1000
        'Else
        '    Form1.Width = 390
        '    'GroupBox3.Visible = False
        'End If



        'SetWindowPos(taskBar, 0&, 0&, 0&, 0&, 0&, SWP_HIDEWINDOW)



        'If File.Exists(Application.StartupPath & "\RS.ini") = False Then
        '    inir.Write("Section1", "RSUSER", "1")
        '    inir.Write("Section1", "RSPASSWORD", "1")
        'Else
        '    RSUser = inir.ReadString("Section1", "RSUSER")
        '    RSPass = inir.ReadString("Section1", "RSPASSWORD")
        'End If
        'Debug.Print("სს " & Chr(34) & "საქკაბელი" & Chr(34))
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Me.Close()
    End Sub
End Class