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
        'giorgi

    End Sub



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