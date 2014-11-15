Public Class Form5
    Public Shared row As Integer
    Dim keluaran As Integer
    Dim id As Integer
    Private Sub Form5_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        id = Form1.ds.Tables("stok").Rows(row)("ID")
        MetroLabel1.Text = Form1.ds.Tables("stok").Rows(row)("nama")
        MetroTextBox13.Text = Form1.ds.Tables("stok").Rows(row)("kod")
        MetroTextBox14.Text = Form1.ds.Tables("stok").Rows(row)("tahun")
        MetroTextBox2.Text = Form1.ds.Tables("stok").Rows(row)("kumpulan")
        MetroTextBox1.Text = Form1.ds.Tables("stok").Rows(row)("unit")
        ComboBox1.Text = Form1.ds.Tables("stok").Rows(row)("lokasi")
        MetroTextBox3.Text = Form1.ds.Tables("stok").Rows(row)("seksyen")
        MetroTextBox4.Text = Form1.ds.Tables("stok").Rows(row)("baris")
        MetroTextBox5.Text = Form1.ds.Tables("stok").Rows(row)("rak")
        MetroTextBox6.Text = Form1.ds.Tables("stok").Rows(row)("tingkat")
        MetroTextBox7.Text = Form1.ds.Tables("stok").Rows(row)("petak")
        MetroTextBox12.Text = Form1.ds.Tables("stok").Rows(row)("maxp")
        MetroTextBox11.Text = Form1.ds.Tables("stok").Rows(row)("menokokp")
        MetroTextBox10.Text = Form1.ds.Tables("stok").Rows(row)("minp")
        MetroTextBox9.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti1t")
        MetroTextBox15.Text = Form1.ds.Tables("stok").Rows(row)("nilai1t")
        MetroTextBox16.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti2t")
        MetroTextBox17.Text = Form1.ds.Tables("stok").Rows(row)("nilai2t")
        MetroTextBox18.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti3t")
        MetroTextBox19.Text = Form1.ds.Tables("stok").Rows(row)("nilai3t")
        MetroTextBox20.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti4t")
        MetroTextBox21.Text = Form1.ds.Tables("stok").Rows(row)("nilai4t")
        MetroTextBox22.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti1k")
        MetroTextBox23.Text = Form1.ds.Tables("stok").Rows(row)("nilai1k")
        MetroTextBox24.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti2k")
        MetroTextBox25.Text = Form1.ds.Tables("stok").Rows(row)("nilai2k")
        MetroTextBox26.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti3k")
        MetroTextBox27.Text = Form1.ds.Tables("stok").Rows(row)("nilai3k")
        MetroTextBox28.Text = Form1.ds.Tables("stok").Rows(row)("kuantiti4k")
        MetroTextBox29.Text = Form1.ds.Tables("stok").Rows(row)("nilai4k")
        MetroTextBox30.Text = Form1.ds.Tables("stok").Rows(row)("qtterimatahunan")
        MetroTextBox31.Text = Form1.ds.Tables("stok").Rows(row)("nilaiterimatahunan")
        MetroTextBox32.Text = Form1.ds.Tables("stok").Rows(row)("qtkeluartahunan")
        MetroTextBox33.Text = Form1.ds.Tables("stok").Rows(row)("nilaikeluartahunan")
    End Sub

    Private Sub MetroButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton8.Click
        keluaran = Convert.ToInt16(MetroTextBox8.Text)
        keluaran = (keluaran / 12)
        MetroTextBox12.Text = (keluaran * 3)
        MetroTextBox11.Text = (keluaran * 2)
        MetroTextBox10.Text = keluaran
    End Sub

    Private Sub MetroButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton2.Click
        Me.Close()
    End Sub

    Private Sub MetroButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton1.Click
        Try
            updatee()
            updatee1()
            MsgBox("Ubah Berjaya")
            Form1.stok(Form1.ListView4)
        Catch ex As Exception
            MsgBox("Ubah Gagal")
        End Try
        

    End Sub
    Public Sub updatee()
        Dim acsconn As System.Data.OleDb.OleDbConnection

        acsconn = New System.Data.OleDb.OleDbConnection
        acsconn.ConnectionString = My.Settings.manageConnectionString
        Dim cmd As New OleDb.OleDbCommand
        Try
            ',nilaikeluartahunan@cc,seksyen=@dd,baris=@ee,rak=@ff,tingkat=@gg,petak=@hh

            acsconn.Open()
            cmd.Connection = acsconn
            cmd.CommandText = "update kwlnstok set nilaikeluartahunan=@a,seksyen=@b,baris=@c,rak=@d,tingkat=@e,petak=@f where ID = @g;"
            cmd.Parameters.AddWithValue("@a", MetroTextBox33.Text)
            cmd.Parameters.AddWithValue("@b", MetroTextBox3.Text)
            cmd.Parameters.AddWithValue("@c", MetroTextBox4.Text)
            cmd.Parameters.AddWithValue("@d", MetroTextBox5.Text)
            cmd.Parameters.AddWithValue("@e", MetroTextBox6.Text)
            cmd.Parameters.AddWithValue("@f", MetroTextBox7.Text)
            cmd.Parameters.AddWithValue("@g", id)
            cmd.ExecuteNonQuery()
            acsconn.Close()

            'MsgBox("berjaya lagi")
        Catch ex As Exception
            MsgBox("gagal ubah")
        End Try
    End Sub
    Public Sub updatee1()
        Dim acsconn As System.Data.OleDb.OleDbConnection

        acsconn = New System.Data.OleDb.OleDbConnection
        acsconn.ConnectionString = My.Settings.manageConnectionString
        Dim cmd As New OleDb.OleDbCommand
        '   Try
        acsconn.Open()
        cmd.Connection = acsconn
        cmd.CommandText = "update kwlnstok set kod=@a,nama=@b,lokasi=@c,unit=@d,kumpulan=@e,tahun=@f,maxp=@g,menokokp=@h,minp=@i,kuantiti1t=@j,nilai1t=@k,kuantiti2t=@l,nilai2t=@m,kuantiti3t=@n,nilai3t=@o,kuantiti4t=@p,nilai4t=@q,kuantiti1k=@r,nilai1k=@s,kuantiti2k=@t,nilai2k=@u,kuantiti3k=@v,nilai3k=@w,kuantiti4k=@x,nilai4k=@y,qtterimatahunan=@z,nilaiterimatahunan=@aa,qtkeluartahunan=@bb where ID = @ii;"
        cmd.Parameters.AddWithValue("@a", MetroTextBox13.Text)
        cmd.Parameters.AddWithValue("@b", MetroLabel1.Text)
        cmd.Parameters.AddWithValue("@c", ComboBox1.Text)
        cmd.Parameters.AddWithValue("@d", MetroTextBox1.Text)
        cmd.Parameters.AddWithValue("@e", MetroTextBox2.Text)
        cmd.Parameters.AddWithValue("@f", MetroTextBox14.Text)
        cmd.Parameters.AddWithValue("@g", MetroTextBox12.Text)
        cmd.Parameters.AddWithValue("@h", MetroTextBox11.Text)
        cmd.Parameters.AddWithValue("@i", MetroTextBox10.Text)
        cmd.Parameters.AddWithValue("@j", MetroTextBox9.Text)
        cmd.Parameters.AddWithValue("@k", MetroTextBox15.Text)
        cmd.Parameters.AddWithValue("@l", MetroTextBox16.Text)
        cmd.Parameters.AddWithValue("@m", MetroTextBox17.Text)
        cmd.Parameters.AddWithValue("@n", MetroTextBox18.Text)
        cmd.Parameters.AddWithValue("@o", MetroTextBox19.Text)
        cmd.Parameters.AddWithValue("@p", MetroTextBox20.Text)
        cmd.Parameters.AddWithValue("@q", MetroTextBox21.Text)
        cmd.Parameters.AddWithValue("@r", MetroTextBox22.Text)
        cmd.Parameters.AddWithValue("@s", MetroTextBox23.Text)
        cmd.Parameters.AddWithValue("@t", MetroTextBox24.Text)
        cmd.Parameters.AddWithValue("@u", MetroTextBox25.Text)
        cmd.Parameters.AddWithValue("@v", MetroTextBox26.Text)
        cmd.Parameters.AddWithValue("@w", MetroTextBox27.Text)
        cmd.Parameters.AddWithValue("@x", MetroTextBox28.Text)
        cmd.Parameters.AddWithValue("@y", MetroTextBox29.Text)
        cmd.Parameters.AddWithValue("@z", MetroTextBox30.Text)
        cmd.Parameters.AddWithValue("@aa", MetroTextBox31.Text)
        cmd.Parameters.AddWithValue("@bb", MetroTextBox32.Text)

        cmd.Parameters.AddWithValue("@ii", id)
        cmd.ExecuteNonQuery()
        acsconn.Close()

        ' MsgBox("Berjaya")
        'Catch ex As Exception
        'MsgBox("gagal ubah")
        'End Try
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        Form11.ShowDialog()
    End Sub
End Class