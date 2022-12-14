Imports System
Imports System.IO
Imports System.Text
Imports System.Reflection

Module Module1
    Public cnnstr As String = "host=172.21.1.200;database=adem_prod5;User ID=adempiere;password=adempiere"
    Public ConStrSQL As String = "Provider=SQLOLEDB;Data Source=10.2.1.150;Initial Catalog=axSAKCABLE;Persist Security Info=True;User ID=smsuser;Password=9"
    Public PSG_Host As String = "172.21.1.200"
    Public PSG_DBName As String = "adem_prod5"
    Public PSG_Port As String = "5432"
    Public PSG_User As String = "adempiere"
    Public PSG_Pass As String = "adempiere"
    Public ADorgID As String
    Public ADorgName As String
    Public ADUserID As String
    Public EmplID As String
    Public WaiBillID As Long
    Public RSUser As String
    Public ApexUser As String
    Public RSPass As String
    Public StartChar As Boolean
    Public CharLenght As Integer
    Public TimeCount As Integer
    Public EmplName As String
    Public FormIsActive As Boolean
    Public DevCountFFL As Integer
    Public Sel1Dev As String
    Public Sel2Dev As String
    Public Sel3Dev As String
    Public BarCode1 As String
    Public BarCode2 As String
    Public BarCode3 As String
    Public KeyCount1 As Integer
    Public KeyCount2 As Integer
    Public KeyCount3 As Integer
    Public IniSet As String
    Public TabNom As Integer
    Public taskBar As Integer
    Public TmapUserID As Double
    Public TmapUserBPartnerID As String
    Public TmapUserName As String
    Public TmapUserRole As String
    Public TmapUserNameEN As String

    Public Dali_income_id_selected As String
    Public Dali_income_ipay_amount_selected As Double
    Public Dali_income_pay_line_selected As String





    Public Function GetPass(ByVal UserName As String, ByRef UserId As Double, ByRef FullName As String, ByRef GroleID As String, ByRef GBPartnerID As String) As String
        Dim mySelectQuery As String = "select user_pass, dali_user_id, user_fullname, _mon_role_id, c_bpartner_id FROM dali_user Where user_name='" & UserName & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetPass = "NULL"

        Try
            While pgReader.Read()
                GetPass = pgReader.GetString(0)
                UserId = pgReader.GetDouble(1)
                FullName = pgReader.GetString(2)
                GroleID = pgReader.GetString(3)
                GBPartnerID = pgReader.GetString(4)
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function CheckCntID(ByVal ChCntID As String, ByRef PBName As String) As String
        Dim tex1 As String
        Dim x1, x2 As Double
        Dim bazac As ADODB.Connection
        bazac = New ADODB.Connection
        bazac.Open(ConStrSQL)
        CheckCntID = "0"
        x1 = 0
        x2 = 0
        Dim bazarec1 As ADODB.Recordset
        bazarec1 = New ADODB.Recordset
        'select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr inner join docs as d on b.Docs_id=d.Docs_id where b.ddate=cast(getdate() as date) and d.Oper_id = '220' group by b.cr,a.acc_nu
        '        tex1 = "select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr where ddate=cast(getdate() as date) group by b.cr,a.acc_nu"
        tex1 = "select cc.cnt_id,c.name from crm.CrmContracts cc, crm.Company c   where cc.cntnum='" & ChCntID & "' and cc.company_id=c.id"
        bazarec1 = bazac.Execute(tex1)
        'bazarec1.MoveFirst()
1:      If bazarec1.EOF Then
            CheckCntID = "0"
            PBName = "NULL"
        Else
            CheckCntID = bazarec1.Fields(0).Value.ToString
            PBName = bazarec1.Fields(1).Value.ToString
        End If
        Return CheckCntID
        bazarec1.Close()
        bazac.Close()
    End Function
    Public Function GetLastID(ByVal TableN As Integer) As Long


        Dim mySelectQuery As String
        mySelectQuery = ""
        Select Case TableN
            Case 1
                mySelectQuery = "SELECT tc.tmap_wire_id FROM tmap_wire tc ORDER BY tc.tmap_wire_id DESC LIMIT 1"
            Case 2
                mySelectQuery = "SELECT tc.tmap_wireline_id FROM tmap_wireline tc ORDER BY tc.tmap_wireline_id DESC LIMIT 1"
            Case 3
                mySelectQuery = "SELECT tc._mon_cable_id FROM _mon_cable tc ORDER BY tc._mon_cable_id DESC LIMIT 1"
            Case 4
                mySelectQuery = "SELECT tc.tmap_cablenorma_id FROM tmap_cablenorma tc ORDER BY tc.tmap_cablenorma_id DESC LIMIT 1"
            Case 5
                mySelectQuery = "SELECT tmap_wiremarka_id FROM tmap_wiremarka ORDER BY tmap_wiremarka_id DESC LIMIT 1"
            Case 6
                mySelectQuery = "SELECT _mon_marka_id FROM _mon_marka ORDER BY _mon_marka_id DESC LIMIT 1"
            Case 7
                mySelectQuery = "SELECT _mon_speed_etalon_id FROM _mon_speed_etalon ORDER BY _mon_speed_etalon_id DESC LIMIT 1"
            Case 8
                mySelectQuery = "SELECT tmap_mat_line_id FROM tmap_mat_line ORDER BY tmap_mat_line_id DESC LIMIT 1"
            Case 9
                mySelectQuery = "SELECT tmap_prod_mat_id FROM tmap_prod_mat ORDER BY tmap_prod_mat_id DESC LIMIT 1"
            Case 10
                mySelectQuery = "SELECT tmap_prod_id FROM tmap_prod ORDER BY tmap_prod_id DESC LIMIT 1" 'tmap_cable_man_line
            Case 11
                mySelectQuery = "SELECT tmap_cable_man_line_id FROM tmap_cable_man_line ORDER BY tmap_cable_man_line_id DESC LIMIT 1" 'tmap_cable_man_line
            Case 12
                mySelectQuery = "SELECT tmap_order_id FROM tmap_order ORDER BY tmap_order_id DESC LIMIT 1" 'tmap_cable_man_line
            Case 13
                mySelectQuery = "SELECT tmap_orderline_id FROM tmap_orderline ORDER BY tmap_orderline_id DESC LIMIT 1" 'tmap_order_processline
            Case 14
                'mySelectQuery = "SELECT tmap_order_processline_id FROM tmap_order_processline ORDER BY tmap_order_processline_id DESC LIMIT 1" 'tmap_order_processline
            Case 15
                mySelectQuery = "SELECT tmap_cable_processline_id FROM tmap_cable_processline ORDER BY tmap_cable_processline_id DESC LIMIT 1" 'tmap_order_processline
            Case 16
                mySelectQuery = "SELECT tmap_cogs_man_history_id FROM tmap_cogs_man_history ORDER BY tmap_cogs_man_history_id DESC LIMIT 1" 'tmap_order_processline
            Case 17
                mySelectQuery = "SELECT tmap_cogs_cost_history_id FROM tmap_cogs_cost_history ORDER BY tmap_cogs_cost_history_id DESC LIMIT 1" 'tmap_order_processline
            Case 18
                mySelectQuery = "SELECT tmap_stock_control_id FROM tmap_stock_control ORDER BY tmap_stock_control_id DESC LIMIT 1" 'tmap_order_processline
                'tmap_stock_tracking_id
            Case 19
                mySelectQuery = "SELECT tmap_stock_tracking_id FROM tmap_stock_tracking ORDER BY tmap_stock_tracking_id DESC LIMIT 1" 'tmap_order_processline
                'tmap_stock_tracking_id  tmap_stock_vendors
            Case 20
                mySelectQuery = "SELECT tmap_stock_vendors_id FROM tmap_stock_vendors ORDER BY tmap_stock_vendors_id DESC LIMIT 1" 'tmap_order_processline
            Case 21
                mySelectQuery = "SELECT tmap_plan_id FROM tmap_plan ORDER BY tmap_plan_id DESC LIMIT 1" 'tmap_order_processline
            Case 22
                mySelectQuery = "SELECT tmap_planline_id FROM tmap_planline ORDER BY tmap_planline_id DESC LIMIT 1" 'tmap_order_processline
            Case 23
                mySelectQuery = "SELECT tmap_tabel_card_id FROM tmap_tabel_card ORDER BY tmap_tabel_card_id DESC LIMIT 1" 'tmap_order_processline
            Case 24
                mySelectQuery = "SELECT tmap_tasks_id FROM tmap_tasks ORDER BY tmap_tasks_id DESC LIMIT 1" 'tmap_order_processline
            Case 25
                mySelectQuery = "SELECT tmap_tasks_line_id FROM tmap_tasks_line ORDER BY tmap_tasks_line_id DESC LIMIT 1" 'tmap_order_processline
            Case 26
                mySelectQuery = "SELECT tmap_orderline_send_id FROM tmap_orderline_send ORDER BY tmap_orderline_send_id DESC LIMIT 1" 'tmap_order_processline
            Case 27
                mySelectQuery = "SELECT hr_tabel_id FROM hr_tabel ORDER BY hr_tabel_id DESC LIMIT 1" 'tmap_order_processline
            Case 28
                mySelectQuery = "SELECT hr_salary_id FROM hr_salary ORDER BY hr_salary_id DESC LIMIT 1" 'tmap_order_processline
            Case 29
                mySelectQuery = "SELECT _mon_sms_id FROM _mon_sms ORDER BY _mon_sms_id DESC LIMIT 1" 'tmap_order_processline
            Case 30
                mySelectQuery = "SELECT hr_danarti_id FROM hr_danarti ORDER BY hr_danarti_id DESC LIMIT 1" 'tmap_order_processline
            Case 31
                mySelectQuery = "SELECT tmap_log_id FROM tmap_log ORDER BY tmap_log_id DESC LIMIT 1" 'tmap_order_processline
            Case 32
                mySelectQuery = "SELECT hr_employee_id FROM hr_employee ORDER BY hr_employee_id DESC LIMIT 1" 'tmap_order_processline
            Case 33
                mySelectQuery = "SELECT c_bpartner_id FROM c_bpartner ORDER BY c_bpartner_id DESC LIMIT 1" 'tmap_order_processline
            Case 34
                mySelectQuery = "SELECT tmap_limiter_id FROM tmap_limiter ORDER BY tmap_limiter_id DESC LIMIT 1" 'tmap_order_processline
        End Select
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetLastID = 0
        Try
            While pgReader.Read()
                GetLastID = pgReader.GetString(0)
            End While
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        Catch myException As Devart.Data.PostgreSql.PgSqlException
            MsgBox(myException.Message)
        End Try
        If GetLastID = 0 Then
            GetLastID = 1000000
        Else
            GetLastID = Val(GetLastID) + 1
        End If
    End Function
    Public Function InsertInLimiter(ByVal ILCompanyID As String, ByVal ILCompanyName As String, ByVal ILAmount As Double, ByVal ILOverdue As String, ByVal ILComment As String) As Boolean
        Dim commandstr As String
        Dim CMLastID As Long

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
        InsertInLimiter = True
        CMLastID = GetLastID(34)
        Dim StDate1 As String = Format(Today, "yyyy-MM-dd") & " " & Format(Now, " HH:mm:ss")

        commandstr = "INSERT INTO tmap_limiter( tmap_limiter_id, recdate, _mon_user_id, company_id, company_name,limit_amount, overdue, status, approve_user, comment, sms_send, sms_id)" & _
        "VALUES (" & CMLastID & ", '" & StDate1 & "', " & TmapUserID & ", " & ILCompanyID & ", '" & ILCompanyName & "', " & ILAmount & ", '" & ILOverdue & "', 'N', '-','" & ILComment & "','N', " & GenSMSID() & ")"
        '"INSERT INTO tmap_cablenorma (tmap_cablenorma_id, tmap_mat_id, weight1km, _mon_user_id, _mon_cable_id) 
        'VALUES (" & InsertNorma() & ", " & TmapMatID & ", " & Weight1km & ", " & TmapUserID & ", " & MonCableID & ");"

        asa2.CommandText = commandstr
        Try
            asa2.ExecuteNonQuery()
        Catch myException As Devart.Data.PostgreSql.PgSqlException
            MsgBox(myException.Message)
            InsertInLimiter = False

        End Try

        asa.Close()
    End Function





    Public Function InsertInTable(ByVal INCommand As String, ByVal IsLoged As Boolean) As String

        InsertInTable = ""

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
        asa2.CommandText = INCommand
        Debug.Print(INCommand)
        Try
            asa2.ExecuteNonQuery()

        Catch myException As Devart.Data.PostgreSql.PgSqlException
            Debug.Print(myException.ErrorCode)
            InsertInTable = myException.ErrorCode
            AddLog(myException.Message)

        End Try

        asa.Close()
    End Function



    Public Function GenSMSID() As Integer
        Dim stdate1 As String = Format(Today, "yyyy-MM-dd")
        Dim stdate2 As String = Format(Today.AddDays(1), "yyyy-MM-dd")
        Dim mySelectQuery As String = "SELECT sms_id FROM tmap_limiter_v where recdate >='" & stdate1 & "' and recdate < '" & stdate2 & "' order by sms_id desc limit 1"
        'SELECT sms_id FROM tmap_limiter_v where recdate >='2021-06-17' and recdate < '2021-06-18' order by sms_id desc limit 1
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GenSMSID = 1
        Try
            While pgReader.Read()
                GenSMSID = pgReader.GetDouble(0) + 1
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try

    End Function
    Public Function GenPassNo() As Integer
        Dim stdate1 As String = Format(Today, "yyyy-MM-dd")
        Dim stdate2 As String = Format(Today.AddDays(1), "yyyy-MM-dd")
        Dim mySelectQuery As String = "SELECT pass_no FROM tmap_pass where pass_date >='" & stdate1 & "' and pass_date < '" & stdate2 & "' order by pass_no desc limit 1"
        'SELECT sms_id FROM tmap_limiter_v where recdate >='2021-06-17' and recdate < '2021-06-18' order by sms_id desc limit 1
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GenPassNo = 1
        Try
            While pgReader.Read()
                GenPassNo = pgReader.GetDouble(0) + 1
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try

    End Function
    Public Function UpdateCntAMT(ByVal InrecordID As String, ByVal INAmt As Double, Indescription As String, ByVal InUserName As String) As Boolean
        Dim tex1 As String
        Dim bazac As ADODB.Connection
        bazac = New ADODB.Connection
        bazac.Open(ConStrSQL)
        UpdateCntAMT = False
        'INSERT INTO crm.[CntOverdraft]   (cnt_id,vg,comment,crtime,cuser) VALUES ('4535','1000','blablabla','2017-01-27 15:30:39.240','smsuser');
        'select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr inner join docs as d on b.Docs_id=d.Docs_id where b.ddate=cast(getdate() as date) and d.Oper_id = '220' group by b.cr,a.acc_nu
        '        tex1 = "select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr where ddate=cast(getdate() as date) group by b.cr,a.acc_nu"
        tex1 = "INSERT INTO crm.[CntOverdraft] (cnt_id,vg,comment,crtime,cuser) VALUES ('" & InrecordID & "','" & mdzime2(Str(INAmt)) & "',N'" & Indescription & "','" & Format(Today, "yyyy-MM-dd") & " " & Format(Now, "hh:mm:ss") & "','" & InUserName & "')"
        Try
            bazac.Execute(tex1)
            UpdateCntAMT = True
        Catch ex As Exception
            UpdateCntAMT = False
        End Try
        bazac.Close()
    End Function
    Public Function CheckCntDate(ByVal ChCntdate As Date) As Boolean
        Dim tex1 As String
        Dim bazac As ADODB.Connection
        bazac = New ADODB.Connection
        bazac.Open(ConStrSQL)
        CheckCntDate = False
        Dim bazarec1 As ADODB.Recordset
        bazarec1 = New ADODB.Recordset
        'select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr inner join docs as d on b.Docs_id=d.Docs_id where b.ddate=cast(getdate() as date) and d.Oper_id = '220' group by b.cr,a.acc_nu
        '        tex1 = "select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr where ddate=cast(getdate() as date) group by b.cr,a.acc_nu"
        tex1 = "SELECT dt FROM Usertbldate where dt='" & Format(ChCntdate, "yyyy-MM-dd") & "'"
        bazarec1 = bazac.Execute(tex1)
        'bazarec1.MoveFirst()
1:      If bazarec1.EOF Then
            CheckCntDate = False
        Else
            CheckCntDate = True
        End If
        Return CheckCntDate
        bazarec1.Close()
        bazac.Close()
    End Function
    Public Function CheckShopUser(ByVal ShopName As String) As Boolean
        Dim tex1 As String
        Dim bazac As ADODB.Connection
        bazac = New ADODB.Connection
        bazac.Open(ConStrSQL)
        CheckShopUser = False
        Dim bazarec1 As ADODB.Recordset
        bazarec1 = New ADODB.Recordset
        'select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr inner join docs as d on b.Docs_id=d.Docs_id where b.ddate=cast(getdate() as date) and d.Oper_id = '220' group by b.cr,a.acc_nu
        '        tex1 = "select a.Acc_nu, sum(o.vg) as vg from orders as o with(nolock) inner join book as b with(nolock) on b.Book_id=o.book_id inner join accounts as a with(nolock) on a.acc=b.cr where ddate=cast(getdate() as date) group by b.cr,a.acc_nu"
        tex1 = "SELECT * FROM [axSAKCABLE].[dbo].[Usertbl] where userid = '" & ShopName & "'"
        bazarec1 = bazac.Execute(tex1)
        'bazarec1.MoveFirst()
1:      If bazarec1.EOF Then
            CheckShopUser = False
        Else
            CheckShopUser = True
        End If
        Return CheckShopUser
        bazarec1.Close()
        bazac.Close()
    End Function
    'Public Function OutAmt(ByVal InUserID As String) As Double
    '    OutAmt = 0
    '    Select Case InUserID
    '        Case "Shop1"
    '            OutAmt = 2000
    '        Case "Shop2"
    '            OutAmt = 2000
    '        Case "Shop4"
    '            OutAmt = 2000
    '        Case "Shop6"
    '            OutAmt = 2000
    '        Case "Shop7"
    '            OutAmt = 20000
    '        Case "Shop8"
    '            OutAmt = 20000
    '        Case "Shop9"
    '            OutAmt = 2000
    '    End Select
    'End Function
    Public Function GetPass(ByVal name As String) As String
        Dim mySelectQuery As String = "select password FROM ad_user au Where au.name='" & name & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetPass = "NULL"

        Try
            While pgReader.Read()
                GetPass = pgReader.GetString(0)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function GetUserID(ByVal name As String) As String
        Dim mySelectQuery As String = "select au.ad_User_id FROM ad_user au Where au.name='" & name & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetUserID = "NULL"

        Try
            While pgReader.Read()
                GetUserID = pgReader.GetString(0)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function GetOrgID(ByVal name As String) As Double
        Dim mySelectQuery As String = "select auo.ad_org_id FROM ad_user au LEFT JOIN ad_user_orgaccess auo ON au.ad_user_id=auo.ad_user_id Where au.name='" & name & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetOrgID = 0

        Try
            While pgReader.Read()
                GetOrgID = pgReader.GetString(0)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function GetOrgname(ByVal GBPID As String, ByRef OrgID As String) As String
        Dim mySelectQuery As String = "SELECT ao.ad_org_id, ao.name FROM ad_org ao Left Join c_bpartner cb ON ao.ad_org_id=cb.ad_org_id Where cb.c_bpartner_id=" & GBPID
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetOrgname = ""

        Try
            While pgReader.Read()
                OrgID = pgReader.GetString(1)
                GetOrgname = pgReader.GetString(1)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function

    Public Function CheckCardID(ByVal CardNo As String) As String
        Dim mySelectQuery As String = "select c_bpartner_id from adempiere.c_bpartner where cardnumber='" & CardNo & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        CheckCardID = ""

        Try
            While pgReader.Read()
                CheckCardID = pgReader.GetString(0)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function GetBPName(ByVal CardNo As String) As String
        On Error GoTo bolo
        Dim mySelectQuery As String = "select Name from adempiere.c_bpartner where cardnumber='" & CardNo & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetBPName = ""


        While pgReader.Read()
            GetBPName = pgReader.GetString(0)
            'Debug.Print(pgReader.GetString(9))
        End While

        ' always call Close when done reading.
        pgReader.Close()
        ' always call Close when done with connection.
        pgConnection.Close()

        Exit Function
bolo:   ShowInfo(Format(TimeOfDay, "HH:mm:ss") & " - " & "თავისუფალი ბარათი")
        GetBPName = ""
    End Function

    Public Function CheckBPCard(ByVal BP_ID As String) As Boolean
        Dim mySelectQuery As String = "select cardnumber from adempiere.c_bpartner where c_bpartner_id=" & BP_ID
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        CheckBPCard = False

        Try
            While pgReader.Read()
                If pgReader.GetString(0) = "1" Then
                    CheckBPCard = True
                Else
                    CheckBPCard = False
                End If

                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function CheckFreeCard(ByVal CardNo As String) As Boolean
        Dim mySelectQuery As String = "select c_bpartner_id from adempiere.c_bpartner where cardnumber='" & CardNo & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        CheckFreeCard = True

        Try
            While pgReader.Read()
                CheckFreeCard = False

                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    'select documentno from adempiere.c_order where C_DocTypeTarget_ID=1000075 Order by dateordered desc limit 1
    Public Function mdzime2(ByVal cifrebi As String) As String
        Dim a, b As String
        Dim x, y As Double
        a = ""
        b = ""
        x = Len(cifrebi)
        For y = 1 To x
            a = Left((Microsoft.VisualBasic.Mid(cifrebi, y)), 1)
            If a = "," Then a = "."
            b = b & a
        Next
        Return b


    End Function
    Public Sub ShowInfo(ByVal InfoText As String)

        'Form1.ListBox1.TopIndex = Form1.ListBox1.SelectedIndex
    End Sub
    Public Function CheckEmplID(ByVal CardNo As String) As String
        Dim mySelectQuery As String = "select c_bpartner_id from adempiere.c_bpartner where cardnumber='" & CardNo & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        CheckEmplID = ""

        Try
            While pgReader.Read()
                CheckEmplID = pgReader.GetString(0)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
        'ShowInfo(CheckEmplID)
    End Function
    Public Function GetEmplName(ByVal EmployID As String) As String
        Dim mySelectQuery As String = "select Name from adempiere.c_bpartner where c_bpartner_id=" & EmployID
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetEmplName = ""


        While pgReader.Read()
            GetEmplName = pgReader.GetString(0)
            'Debug.Print(pgReader.GetString(9))
        End While

        ' always call Close when done reading.
        pgReader.Close()
        ' always call Close when done with connection.
        pgConnection.Close()
    End Function

    Public Function GetTime() As String
        Dim mySelectQuery As String = "SELECT CURRENT_TIME"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetTime = 0
        Try
            While pgReader.Read()
                GetTime = pgReader.GetString(0)
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
        GetTime = Microsoft.VisualBasic.Left(GetTime, 8)
    End Function
    Public Function GetDate() As String
        Dim mySelectQuery As String = "SELECT CURRENT_TIMESTAMP"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetDate = 0
        Try
            While pgReader.Read()
                GetDate = pgReader.GetString(0)
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
        GetDate = Format(Today, "yyyy-MM-dd") 'Microsoft.VisualBasic.Left(GetDate, 10)
    End Function
    Public Sub AddLog(ByVal LogText As String)
        Dim AllLog As String
        Dim filename As String
        filename = Format(Today, "yyyy-MM") & ".log"
        AllLog = Format(Today, "dd.MM.yyyy") & " " & Format(Now, "HH:mm:ss") & ";" & LogText & vbCrLf
        'If File.Exists(Application.StartupPath & "\" & filename) = False Then
        '    PictureBox2.Load(Application.StartupPath & "\pics\noimg.jpeg")
        'Else
        '    PictureBox2.Load(Application.StartupPath & "\pics\" & CardPIN & ".jpg")
        'End If
        Try
            Using myWriteFile As New IO.StreamWriter(Application.StartupPath & "\" & filename, True, System.Text.Encoding.UTF8)
                myWriteFile.Write(AllLog)
            End Using
        Catch ex As IO.IOException
            MsgBox(ex.ToString) 'ошибка
        End Try


    End Sub

    Public Function mdzime3(ByVal cifrebi As String) As String
        Dim a, b As String
        Dim x, y As Double
        a = ""
        b = ""
        x = Len(cifrebi)
        For y = 1 To x
            a = Left((Microsoft.VisualBasic.Mid(cifrebi, y)), 1)
            If a = "." Then a = ","
            b = b & a
        Next
        Return b
    End Function
    Public Function GetBPIDFromPIN(ByVal PIN As String, ByRef BP_Name As String) As String
        Dim mySelectQuery As String = "SELECT bp.c_bpartner_id, bp.name FROM adempiere.c_bpartner bp where isemployee='Y' and isshop_employee='Y' and ad_client_id=1000000 and cardnumber='" & PIN & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetBPIDFromPIN = 0

        Try
            While pgReader.Read()
                GetBPIDFromPIN = pgReader.GetString(0)
                BP_Name = pgReader.GetString(1)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function GetDuration(ByVal TRecID As String) As String
        Dim mySelectQuery As String = "SELECT duration FROM adempiere._saq_employee_timecard where _saq_employee_timecard_id=" & TRecID
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetDuration = ""

        Try
            While pgReader.Read()
                GetDuration = pgReader.GetString(0)
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function CheckOutTime(ByVal BP_ID As String) As Boolean
        Dim mySelectQuery As String = "SELECT tc.c_bpartner_id FROM _saq_employee_timecard tc where tc.outtime is null and  tc.c_bpartner_id=" & BP_ID & " order by tc.intime desc limit 1"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        CheckOutTime = True

        Try
            While pgReader.Read()
                CheckOutTime = False
                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function GetLastID() As Long
        Dim mySelectQuery As String = "SELECT tc._saq_employee_timecard_id FROM _saq_employee_timecard tc ORDER BY tc._saq_employee_timecard_id DESC LIMIT 1"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        GetLastID = 0
        Try
            While pgReader.Read()
                GetLastID = pgReader.GetString(0)
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
        If GetLastID = 0 Then


            GetLastID = 1000000
        Else
            GetLastID = Val(GetLastID) + 1
        End If
    End Function
    Public Function CheckInTime(ByVal BP_ID As String) As String

        Dim mySelectQuery As String = "SELECT tc._saq_employee_timecard_id,tc.outtime FROM _saq_employee_timecard tc where tc.c_bpartner_id=" & BP_ID & " and tc.outtime is null limit 1"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        Dim RecID As String
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        CheckInTime = "1"

        Try
            While pgReader.Read()
                RecID = pgReader.GetString(0)
                If pgReader.GetString(1) = "" Then
                    CheckInTime = RecID
                Else
                    CheckInTime = "0"
                End If
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function
    Public Function InsertInTable(ByVal INCommand As String) As String
        InsertInTable = ""
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
        asa2.CommandText = INCommand
        Debug.Print(INCommand)
        Try
            asa2.ExecuteNonQuery()
        Catch myException As Devart.Data.PostgreSql.PgSqlException
            Debug.Print(myException.ErrorCode)
            InsertInTable = myException.ErrorCode
            AddLog(myException.Message)

        End Try
        asa.Close()
    End Function







    Public Sub FillFuelCombo1(ByVal ZComboControl As ComboBox)
        ZComboControl.ValueMember = (Nothing)
        ZComboControl.DisplayMember = (Nothing)
        ZComboControl.DataSource = (Nothing)
        Dim asa As Devart.Data.PostgreSql.PgSqlConnection
        asa = New Devart.Data.PostgreSql.PgSqlConnection
        asa.Unicode = True
        asa.Host = PSG_Host
        asa.Database = PSG_DBName
        asa.Port = PSG_Port
        asa.UserId = PSG_User
        asa.Password = PSG_Pass
        asa.Open()
        Dim asa1 As New Devart.Data.PostgreSql.PgSqlDataAdapter("select company_name as company_name From dali_company  order by company_name", asa)

        Dim ds1 As New DataSet
        Dim ds2 As New DataSet
        Dim ds3 As New DataSet
        Dim ds4 As New DataSet
        Dim ds5 As New DataSet
        Dim ds6 As New DataSet
        asa1.Fill(ds1, "dali_company")
        asa.Close()
        ZComboControl.DataSource = ds1.Tables("dali_company")
        ZComboControl.DisplayMember = "company_name"
        ZComboControl.ValueMember = "company_name"
        ZComboControl.Update()
    End Sub






    Public Sub FillFuelCombo2(ByVal InCombo As ComboBox)
        Dim mySelectQuery As String = "select company_name as company_name From dali_company  order by company_name"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        InCombo.Items.Clear()
        InCombo.Items.Add("")

        Try
            While pgReader.Read()
                InCombo.Items.Add(pgReader.GetString(0))

            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Sub



    Public Function Get_DCOMPANY_ID_From_name(ByVal GCableName As String) As String
        Dim mySelectQuery As String = "select dali_company_id From dali_company where company_name ='" & GCableName & "'"
        Dim pgConnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        Dim pgCommand As New Devart.Data.PostgreSql.PgSqlCommand(mySelectQuery, pgConnection)
        pgConnection.Unicode = True
        pgConnection.Open()
        Dim pgReader As Devart.Data.PostgreSql.PgSqlDataReader = pgCommand.ExecuteReader()
        Get_DCOMPANY_ID_From_name = "0"

        Try
            While pgReader.Read()
                Get_DCOMPANY_ID_From_name = pgReader.GetString(0)

                'Debug.Print(pgReader.GetString(9))
            End While
        Finally
            ' always call Close when done reading.
            pgReader.Close()
            ' always call Close when done with connection.
            pgConnection.Close()
        End Try
    End Function

    Public Function FillPSQLDataSet(ByVal QueryStr As String) As DataSet
        Dim asa As New Devart.Data.PostgreSql.PgSqlConnection
        Dim CFQueryText As String
        asa.Unicode = True
        asa.Host = PSG_Host
        asa.Database = PSG_DBName
        asa.Port = PSG_Port
        asa.UserId = PSG_User
        asa.Password = PSG_Pass
        asa.Open()
        Dim ds As New DataSet()
        CFQueryText = ""
        Dim asa3 As New Devart.Data.PostgreSql.PgSqlDataAdapter(QueryStr, asa)
        asa3.Fill(ds, "Authors_table")
        asa.Close()
        FillPSQLDataSet = ds
    End Function


    Public Function FillPSQLDataSet2(ByVal QueryStr As String) As DataSet
        Dim connectionString As String = ConStrSQL
        Dim pconnection As New Devart.Data.PostgreSql.PgSqlConnection(cnnstr)
        pconnection.Unicode = True
        pconnection.ConnectionTimeout = 300
        pconnection.Open()

        Dim cmd As New Devart.Data.PostgreSql.PgSqlCommand
        cmd.CommandTimeout = 300
        cmd.CommandText = QueryStr
        cmd.Connection = pconnection
        Dim padapter As New Devart.Data.PostgreSql.PgSqlDataAdapter(cmd)
        padapter.SelectCommand = cmd

        Dim ds As New DataSet()
        padapter.Fill(ds)
        pconnection.Close()
        FillPSQLDataSet2 = ds
    End Function




    Public Sub EnableDoubleBuffered(ByVal dgv As DataGridView)

        Dim dgvType As Type = dgv.[GetType]()

        Dim pi As PropertyInfo = dgvType.GetProperty("DoubleBuffered", _
                                                     BindingFlags.Instance Or BindingFlags.NonPublic)

        pi.SetValue(dgv, True, Nothing)

    End Sub


End Module
