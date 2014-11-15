Public Class Form12
    Public Shared row As Integer
    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MetroTextBox42.Text = Form1.ds.Tables("ked").Rows(row).Item("kod")
        MetroTextBox1.Text = Form1.ds.Tables("ked").Rows(row).Item("sa1")
        MetroTextBox2.Text = Form1.ds.Tables("ked").Rows(row).Item("sa11")
        MetroTextBox3.Text = Form1.ds.Tables("ked").Rows(row).Item("sa2")
        MetroTextBox4.Text = Form1.ds.Tables("ked").Rows(row).Item("sa22")
        MetroTextBox5.Text = Form1.ds.Tables("ked").Rows(row).Item("sa3")
        MetroTextBox6.Text = Form1.ds.Tables("ked").Rows(row).Item("sa33")
        MetroTextBox7.Text = Form1.ds.Tables("ked").Rows(row).Item("sa4")
        MetroTextBox8.Text = Form1.ds.Tables("ked").Rows(row).Item("sa44")
        MetroTextBox9.Text = Form1.ds.Tables("ked").Rows(row).Item("b1")
        MetroTextBox10.Text = Form1.ds.Tables("ked").Rows(row).Item("b11")
        MetroTextBox11.Text = Form1.ds.Tables("ked").Rows(row).Item("b2")
        MetroTextBox12.Text = Form1.ds.Tables("ked").Rows(row).Item("b22")
        MetroTextBox13.Text = Form1.ds.Tables("ked").Rows(row).Item("b3")
        MetroTextBox14.Text = Form1.ds.Tables("ked").Rows(row).Item("b33")
        MetroTextBox15.Text = Form1.ds.Tables("ked").Rows(row).Item("b4")
        MetroTextBox16.Text = Form1.ds.Tables("ked").Rows(row).Item("b44")
        MetroTextBox17.Text = Form1.ds.Tables("ked").Rows(row).Item("k1")
        MetroTextBox18.Text = Form1.ds.Tables("ked").Rows(row).Item("k11")
        MetroTextBox19.Text = Form1.ds.Tables("ked").Rows(row).Item("k2")
        MetroTextBox20.Text = Form1.ds.Tables("ked").Rows(row).Item("k22")
        MetroTextBox21.Text = Form1.ds.Tables("ked").Rows(row).Item("k3")
        MetroTextBox22.Text = Form1.ds.Tables("ked").Rows(row).Item("k33")
        MetroTextBox23.Text = Form1.ds.Tables("ked").Rows(row).Item("k4")
        MetroTextBox24.Text = Form1.ds.Tables("ked").Rows(row).Item("k44")
        MetroTextBox25.Text = Form1.ds.Tables("ked").Rows(row).Item("ss1")
        MetroTextBox26.Text = Form1.ds.Tables("ked").Rows(row).Item("ss11")
        MetroTextBox27.Text = Form1.ds.Tables("ked").Rows(row).Item("ss2")
        MetroTextBox28.Text = Form1.ds.Tables("ked").Rows(row).Item("ss22")
        MetroTextBox29.Text = Form1.ds.Tables("ked").Rows(row).Item("ss3")
        MetroTextBox30.Text = Form1.ds.Tables("ked").Rows(row).Item("ss33")
        MetroTextBox31.Text = Form1.ds.Tables("ked").Rows(row).Item("ss4")
        MetroTextBox32.Text = Form1.ds.Tables("ked").Rows(row).Item("ss44")
        MetroTextBox33.Text = Form1.ds.Tables("ked").Rows(row).Item("kps1")
        MetroTextBox34.Text = Form1.ds.Tables("ked").Rows(row).Item("kps2")
        MetroTextBox35.Text = Form1.ds.Tables("ked").Rows(row).Item("kps3")
        MetroTextBox36.Text = Form1.ds.Tables("ked").Rows(row).Item("kps4")
        MetroTextBox37.Text = Form1.ds.Tables("ked").Rows(row).Item("ntb1")
        MetroTextBox38.Text = Form1.ds.Tables("ked").Rows(row).Item("ntb11")
        MetroTextBox39.Text = Form1.ds.Tables("ked").Rows(row).Item("ntk1")
        MetroTextBox40.Text = Form1.ds.Tables("ked").Rows(row).Item("ntk11")
        MetroTextBox41.Text = Form1.ds.Tables("ked").Rows(row).Item("kps")


    End Sub

    Private Sub MetroButton13_Click(sender As Object, e As EventArgs) Handles MetroButton13.Click
        Me.Close()
    End Sub
End Class