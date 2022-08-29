Public Class PAY

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Debug.Print("wavida gadaxda")
        AddLog("wavida gadaxda")
        ' Dali_income_id_selected = 1
        Dim PeriodStr As String = DateTimePicker5.Value.Year.ToString & IIf(Len(DateTimePicker5.Value.Month.ToString) = 1, "0" & DateTimePicker5.Value.Month.ToString, DateTimePicker5.Value.Month.ToString)


        Debug.Print(PeriodStr)

        AddLog(PeriodStr)
        Dim stdate1 As String = Format(Today, "yyyy-MM-dd") & " " & Format(Now, " HH:mm:ss")


        Debug.Print(stdate1)
        AddLog(stdate1)
        Dim inserttext As String

        inserttext = "INSERT INTO dali_income_line( dali_income_id, period,  isactive,ipay_amount,ipay_day)   VALUES ('" & Dali_income_id_selected & "', '" & PeriodStr & "', 'Y','" & TextBox2.Text & "','" & Format(DateTimePickerPD.Value, "yyyy-MM-dd") & "');"
        Debug.Print(inserttext)
        AddLog(inserttext)
        InsertInTable(inserttext, False)

        Me.Close()

    End Sub

    Private Sub PAY_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Text = Dali_income_ipay_amount_selected
    End Sub
End Class