Public Class Form7
    Dim ds As New DataSet
    Dim row As Integer
    Dim b As Boolean = False
    ' Dim keluaran As Integer
    Private Sub MetroButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton3.Click
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Try
            oleConn.Open()
            da = New OleDb.OleDbDataAdapter("select kod,nama,tarikh from detail where kod like '%" & MetroTextBox34.Text & "%'", oleConn)
            da.Fill(ds, "infox")
            oleConn.Close()

            MetroLabel1.Text = ds.Tables("infox").Rows(row)("nama")
            MetroTextBox13.Text = ds.Tables("infox").Rows(row)("kod")
            '  parasstok()
            terimaan()
            keluaran()
            stoktahunan()
        Catch ex As Exception
            MsgBox("Tiada Data Dalam Database.")
        End Try
    End Sub

    Private Sub MetroButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton1.Click
        If MetroTextBox2.Text = "" Or MetroTextBox1.Text = "" Or MetroTextBox14.Text = "" Or MetroTextBox3.Text = "" Or MetroTextBox4.Text = "" Or MetroTextBox5.Text = "" Or MetroTextBox6.Text = "" Or MetroTextBox7.Text = "" Then
            MsgBox("Sila Isi Ruangan Terlebih Dahulu")
            b = False
        Else
            Try
                insert1()
                ' insert2()
                MsgBox("Berjaya Ubah")
                b = True
                'Form1.stok(Form1.ListView4)
            Catch ex As Exception
                MsgBox("gagal ubah")
            End Try
           
        End If
       
       
    End Sub
    Public Sub insert1()
        Dim acsconn As System.Data.OleDb.OleDbConnection

        acsconn = New System.Data.OleDb.OleDbConnection
        acsconn.ConnectionString = My.Settings.manageConnectionString
        Dim cmd As New OleDb.OleDbCommand
        Try
            acsconn.Open()
            cmd.Connection = acsconn
            cmd.CommandText = "insert into kwlnstok(kod,nama,lokasi,unit,kumpulan,tahun,maxp,menokokp,minp,kuantiti1t,nilai1t,kuantiti2t,nilai2t,kuantiti3t,nilai3t,kuantiti4t,nilai4t,kuantiti1k,nilai1k,kuantiti2k,nilai2k,kuantiti3k,nilai3k,kuantiti4k,nilai4k,qtterimatahunan,nilaiterimatahunan,qtkeluartahunan,nilaikeluartahunan,seksyen,baris,rak,tingkat,petak) values(@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n,@o,@p,@q,@r,@s,@t,@u,@v,@w,@x,@y,@z,@aa,@bb,@cc,@dd,@ee,@ff,@gg,@hh);"
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
            cmd.Parameters.AddWithValue("@cc", MetroTextBox33.Text)
            cmd.Parameters.AddWithValue("@dd", MetroTextBox3.Text)
            cmd.Parameters.AddWithValue("@ee", MetroTextBox4.Text)
            cmd.Parameters.AddWithValue("@ff", MetroTextBox5.Text)
            cmd.Parameters.AddWithValue("@gg", MetroTextBox6.Text)
            cmd.Parameters.AddWithValue("@hh", MetroTextBox7.Text)
            ' cmd.Parameters.AddWithValue("@ii", id)
            cmd.ExecuteNonQuery()
            acsconn.Close()

            'MsgBox("Berjaya")
        Catch ex As Exception
            MsgBox("gagal ubah 1")
        End Try
    End Sub
    Public Sub insert2()
        Dim acsconn As System.Data.OleDb.OleDbConnection
        acsconn = New System.Data.OleDb.OleDbConnection
        acsconn.ConnectionString = My.Settings.manageConnectionString
        Dim cmd As New OleDb.OleDbCommand
        Try
            acsconn.Open()
            cmd.Connection = acsconn
            cmd.CommandText = "insert into kwlnstok(nilaikeluartahunan,seksyen,baris,rak,tingkat,petak) values(@a,@b,@c,@d,@e,@f);"
            cmd.Parameters.AddWithValue("@a", MetroTextBox33.Text)
            cmd.Parameters.AddWithValue("@b", MetroTextBox3.Text)
            cmd.Parameters.AddWithValue("@c", MetroTextBox4.Text)
            cmd.Parameters.AddWithValue("@d", MetroTextBox5.Text)
            cmd.Parameters.AddWithValue("@e", MetroTextBox6.Text)
            cmd.Parameters.AddWithValue("@f", MetroTextBox7.Text)
            'cmd.Parameters.AddWithValue("@g", id)
            cmd.ExecuteNonQuery()
            acsconn.Close()

            'MsgBox("berjaya lagi")
        Catch ex As Exception
            MsgBox("gagal ubah 2")
        End Try
    End Sub

    Private Sub MetroButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton2.Click
        b = False
        Form1.stok(Form1.ListView4)
        Me.Close()
    End Sub
    Public Sub parasstok()
        
    End Sub
    Public Sub terimaan()
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Try
            oleConn.Open()
            da = New OleDb.OleDbDataAdapter("select kod,nama,harga,kuantiti,jumlah,tarikh,suku from tambah where kod like '%" & MetroTextBox34.Text & "%'", oleConn)
            da.Fill(ds, "terima")
            oleConn.Close()
        Catch ex As Exception
            MsgBox("Tiada Data Dalam Database.")
        End Try
        Dim kuantitipertama As Decimal = 0
        Dim nilaipertama As Decimal = 0
        Dim kuantitikedua As Decimal = 0
        Dim nilaikedua As Decimal = 0
        Dim kuantitiketiga As Decimal = 0
        Dim nilaiketiga As Decimal = 0
        Dim kuantitikeempat As Decimal = 0
        Dim nilaikeempat As Decimal = 0
        For row As Integer = 0 To ds.Tables("terima").Rows.Count - 1
            If ds.Tables("terima").Rows(row).Item("suku") = 1 Then

                kuantitipertama += ds.Tables("terima").Rows(row).Item("kuantiti")
                nilaipertama += ds.Tables("terima").Rows(row).Item("jumlah")

            ElseIf ds.Tables("terima").Rows(row).Item("suku") = 2 Then

                kuantitikedua += ds.Tables("terima").Rows(row).Item("kuantiti")
                nilaikedua += ds.Tables("terima").Rows(row).Item("jumlah")

            ElseIf ds.Tables("terima").Rows(row).Item("suku") = 3 Then

                kuantitiketiga += ds.Tables("terima").Rows(row).Item("kuantiti")
                nilaiketiga += ds.Tables("terima").Rows(row).Item("jumlah")

            ElseIf ds.Tables("terima").Rows(row).Item("suku") = 4 Then

                kuantitikeempat += ds.Tables("terima").Rows(row).Item("kuantiti")
                nilaikeempat += ds.Tables("terima").Rows(row).Item("jumlah")

            End If
        Next
        MetroTextBox9.Text = kuantitipertama
        MetroTextBox15.Text = nilaipertama.ToString("N2")
        MetroTextBox16.Text = kuantitikedua
        MetroTextBox17.Text = nilaikedua.ToString("N2")
        MetroTextBox18.Text = kuantitiketiga
        MetroTextBox19.Text = nilaiketiga.ToString("N2")
        MetroTextBox20.Text = kuantitikeempat
        MetroTextBox21.Text = nilaikeempat.ToString("N2")
    End Sub
    Public Sub keluaran()
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Try
            oleConn.Open()
            da = New OleDb.OleDbDataAdapter("select kod,nama,harga,kuantiti,jumlah,tarikh,suku from kurang where kod like '%" & MetroTextBox34.Text & "%'", oleConn)
            da.Fill(ds, "keluar")
            oleConn.Close()
        Catch ex As Exception
            MsgBox("Tiada Data Dalam Database.")
        End Try
        Dim kuantitipertama As Decimal = 0
        Dim nilaipertama As Decimal = 0
        Dim kuantitikedua As Decimal = 0
        Dim nilaikedua As Decimal = 0
        Dim kuantitiketiga As Decimal = 0
        Dim nilaiketiga As Decimal = 0
        Dim kuantitikeempat As Decimal = 0
        Dim nilaikeempat As Decimal = 0
        For row As Integer = 0 To ds.Tables("keluar").Rows.Count - 1
            If ds.Tables("keluar").Rows(row).Item("suku") = 1 Then

                kuantitipertama += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaipertama += ds.Tables("keluar").Rows(row).Item("jumlah")

            ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 2 Then

                kuantitikedua += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaikedua += ds.Tables("keluar").Rows(row).Item("jumlah")

            ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 3 Then

                kuantitiketiga += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaiketiga += ds.Tables("keluar").Rows(row).Item("jumlah")

            ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 4 Then

                kuantitikeempat += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaikeempat += ds.Tables("keluar").Rows(row).Item("jumlah")

            End If
        Next
        MetroTextBox22.Text = kuantitipertama
        MetroTextBox23.Text = nilaipertama.ToString("N2")
        MetroTextBox24.Text = kuantitikedua
        MetroTextBox25.Text = nilaikedua.ToString("N2")
        MetroTextBox26.Text = kuantitiketiga
        MetroTextBox27.Text = nilaiketiga.ToString("N2")
        MetroTextBox28.Text = kuantitikeempat
        MetroTextBox29.Text = nilaikeempat.ToString("N2")
    End Sub
    Public Sub stoktahunan()
        Dim kuantititerima As Decimal = 0
        Dim nilaiterima As Decimal = 0
        Dim kuantitikeluar As Decimal = 0
        Dim nilaikeluar As Decimal = 0
        Dim kuantitipertama As Decimal = 0
        Dim nilaipertama As Decimal = 0
        Dim kuantitikedua As Decimal = 0
        Dim nilaikedua As Decimal = 0
        Dim kuantitiketiga As Decimal = 0
        Dim nilaiketiga As Decimal = 0
        Dim kuantitikeempat As Decimal = 0
        Dim nilaikeempat As Decimal = 0
        kuantitipertama = Convert.ToDecimal(MetroTextBox9.Text)
        kuantitikedua = Convert.ToDecimal(MetroTextBox16.Text)
        kuantitiketiga = Convert.ToDecimal(MetroTextBox18.Text)
        kuantitikeempat = Convert.ToDecimal(MetroTextBox20.Text)
        nilaipertama = Convert.ToDecimal(MetroTextBox15.Text)
        nilaikedua = Convert.ToDecimal(MetroTextBox17.Text)
        nilaiketiga = Convert.ToDecimal(MetroTextBox19.Text)
        nilaikeempat = Convert.ToDecimal(MetroTextBox21.Text)
        kuantititerima = kuantitipertama + kuantitikedua + kuantitiketiga + kuantitikeempat
        nilaiterima = nilaipertama + nilaikedua + nilaiketiga + nilaikeempat

        MetroTextBox30.Text = kuantititerima
        MetroTextBox31.Text = nilaiterima.ToString("N2")

        kuantitipertama = Convert.ToDecimal(MetroTextBox22.Text)
        kuantitikedua = Convert.ToDecimal(MetroTextBox24.Text)
        kuantitiketiga = Convert.ToDecimal(MetroTextBox26.Text)
        kuantitikeempat = Convert.ToDecimal(MetroTextBox28.Text)
        nilaipertama = Convert.ToDecimal(MetroTextBox23.Text)
        nilaikedua = Convert.ToDecimal(MetroTextBox25.Text)
        nilaiketiga = Convert.ToDecimal(MetroTextBox27.Text)
        nilaikeempat = Convert.ToDecimal(MetroTextBox29.Text)
        kuantititerima = kuantitipertama + kuantitikedua + kuantitiketiga + kuantitikeempat
        nilaiterima = nilaipertama + nilaikedua + nilaiketiga + nilaikeempat

        MetroTextBox32.Text = kuantititerima
        MetroTextBox33.Text = nilaiterima.ToString("N2")
        Dim jumlah As Decimal = 0
        'For row As Integer = 0 To ds.Tables("infox").Rows.Count - 1
        '    jumlah += ds.Tables("infox").Rows(row).Item("kuantiti")
        'Next
        ' keluaran = Convert.ToInt16(MetroTextBox8.Text)
        'keluaran = (keluaran / 12)
        jumlah = kuantititerima
        jumlah = (jumlah / 12)
        MetroTextBox12.Text = (jumlah * 3)
        MetroTextBox11.Text = (jumlah * 2)
        MetroTextBox10.Text = jumlah
    End Sub

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        If b Then
            Form10.ShowDialog()
        Else
            MsgBox("Sila Simpan Data Terlebih Dahulu.")
        End If

    End Sub

   
End Class