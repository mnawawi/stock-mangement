Public Class Form9
    Dim ds As New DataSet
    Dim row As Integer
    Dim thisdate As Date
    Dim thisyear As Integer

    Dim jumlah As Decimal = 0
    Dim kuantitipertama As Decimal = 0
    Dim nilaipertama As Decimal = 0
    Dim kuantitikedua As Decimal = 0
    Dim nilaikedua As Decimal = 0
    Dim kuantitiketiga As Decimal = 0
    Dim nilaiketiga As Decimal = 0
    Dim kuantitikeempat As Decimal = 0
    Dim nilaikeempat As Decimal = 0
    Dim kuantitipertama1 As Decimal = 0
    Dim nilaipertama1 As Decimal = 0
    Dim kuantitikedua1 As Decimal = 0
    Dim nilaikedua1 As Decimal = 0
    Dim kuantitiketiga1 As Decimal = 0
    Dim nilaiketiga1 As Decimal = 0
    Dim kuantitikeempat1 As Decimal = 0
    Dim nilaikeempat1 As Decimal = 0

    Dim sedia As Decimal = 0
    Dim sedia1 As Decimal = 0
    Dim kadar As Decimal = 0
    Dim kadar1 As Decimal = 0
    Dim kadar2 As Decimal = 0
    Dim kadar3 As Decimal = 0
    Dim jumlahkadar As Decimal = 0

    Dim totalkuantiti As Decimal = 0
    Dim totalnilai As Decimal = 0
    Dim totalkuantiti1 As Decimal = 0
    Dim totalnilai1 As Decimal = 0
    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    'Public Sub terimaan()
    '    Dim oleConn As System.Data.OleDb.OleDbConnection
    '    Dim da As OleDb.OleDbDataAdapter
    '    ds = New DataSet
    '    oleConn = New System.Data.OleDb.OleDbConnection
    '    oleConn.ConnectionString = My.Settings.manageConnectionString
    '    Try
    '        oleConn.Open()
    '        da = New OleDb.OleDbDataAdapter("select kod,nama,harga,kuantiti,jumlah,tarikh,suku from tambah where kod like '%" & MetroTextBox42.Text & "%'", oleConn)
    '        da.Fill(ds, "terima")
    '        oleConn.Close()
    '    Catch ex As Exception
    '        MsgBox("Tiada Data Dalam Database.")
    '    End Try
    '    'Dim kuantitipertama As Decimal = 0
    '    'Dim nilaipertama As Decimal = 0
    '    'Dim kuantitikedua As Decimal = 0
    '    'Dim nilaikedua As Decimal = 0
    '    'Dim kuantitiketiga As Decimal = 0
    '    'Dim nilaiketiga As Decimal = 0
    '    'Dim kuantitikeempat As Decimal = 0
    '    'Dim nilaikeempat As Decimal = 0
    '    For row As Integer = 0 To ds.Tables("terima").Rows.Count - 1
    '        If ds.Tables("terima").Rows(row).Item("suku") = 1 Then

    '            kuantitipertama += ds.Tables("terima").Rows(row).Item("kuantiti")
    '            nilaipertama += ds.Tables("terima").Rows(row).Item("jumlah")

    '        ElseIf ds.Tables("terima").Rows(row).Item("suku") = 2 Then

    '            kuantitikedua += ds.Tables("terima").Rows(row).Item("kuantiti")
    '            nilaikedua += ds.Tables("terima").Rows(row).Item("jumlah")

    '        ElseIf ds.Tables("terima").Rows(row).Item("suku") = 3 Then

    '            kuantitiketiga += ds.Tables("terima").Rows(row).Item("kuantiti")
    '            nilaiketiga += ds.Tables("terima").Rows(row).Item("jumlah")

    '        ElseIf ds.Tables("terima").Rows(row).Item("suku") = 4 Then

    '            kuantitikeempat += ds.Tables("terima").Rows(row).Item("kuantiti")
    '            nilaikeempat += ds.Tables("terima").Rows(row).Item("jumlah")

    '        End If
    '    Next
    '    MetroTextBox9.Text = kuantitipertama
    '    MetroTextBox10.Text = nilaipertama.ToString("N2")
    '    MetroTextBox11.Text = kuantitikedua
    '    MetroTextBox12.Text = nilaikedua.ToString("N2")
    '    MetroTextBox13.Text = kuantitiketiga
    '    MetroTextBox14.Text = nilaiketiga.ToString("N2")
    '    MetroTextBox15.Text = kuantitikeempat
    '    MetroTextBox16.Text = nilaikeempat.ToString("N2")
    'End Sub
    Public Sub terimaan()
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Try
            oleConn.Open()
            da = New OleDb.OleDbDataAdapter("select kod,nama,harga,kuantiti,jumlah,tarikh,suku from tambah", oleConn)
            da.Fill(ds, "terima")
            oleConn.Close()
        Catch ex As Exception
            MsgBox("Tiada Data Dalam Database.")
        End Try
        'Dim kuantitipertama As Decimal = 0
        'Dim nilaipertama As Decimal = 0
        'Dim kuantitikedua As Decimal = 0
        'Dim nilaikedua As Decimal = 0
        'Dim kuantitiketiga As Decimal = 0
        'Dim nilaiketiga As Decimal = 0
        'Dim kuantitikeempat As Decimal = 0
        'Dim nilaikeempat As Decimal = 0
        thisdate = ds.Tables("terima").Rows(row).Item("tarikh")
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
        MetroTextBox10.Text = nilaipertama.ToString("N2")
        MetroTextBox11.Text = kuantitikedua
        MetroTextBox12.Text = nilaikedua.ToString("N2")
        MetroTextBox13.Text = kuantitiketiga
        MetroTextBox14.Text = nilaiketiga.ToString("N2")
        MetroTextBox15.Text = kuantitikeempat
        MetroTextBox16.Text = nilaikeempat.ToString("N2")
    End Sub
    Public Sub keluaran()
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Try
            oleConn.Open()
            da = New OleDb.OleDbDataAdapter("select kod,nama,harga,kuantiti,jumlah,tarikh,suku from kurang", oleConn)
            da.Fill(ds, "keluar")
            oleConn.Close()
        Catch ex As Exception
            MsgBox("Tiada Data Dalam Database.")
        End Try
        'Dim kuantitipertama As Decimal = 0
        'Dim nilaipertama As Decimal = 0
        'Dim kuantitikedua As Decimal = 0
        'Dim nilaikedua As Decimal = 0
        'Dim kuantitiketiga As Decimal = 0
        'Dim nilaiketiga As Decimal = 0
        'Dim kuantitikeempat As Decimal = 0
        'Dim nilaikeempat As Decimal = 0
        For row As Integer = 0 To ds.Tables("keluar").Rows.Count - 1
            If ds.Tables("keluar").Rows(row).Item("suku") = 1 Then

                kuantitipertama1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaipertama1 += ds.Tables("keluar").Rows(row).Item("jumlah")

            ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 2 Then

                kuantitikedua1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaikedua1 += ds.Tables("keluar").Rows(row).Item("jumlah")

            ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 3 Then

                kuantitiketiga1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaiketiga1 += ds.Tables("keluar").Rows(row).Item("jumlah")

            ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 4 Then

                kuantitikeempat1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
                nilaikeempat1 += ds.Tables("keluar").Rows(row).Item("jumlah")

            End If
        Next
        MetroTextBox17.Text = kuantitipertama1
        MetroTextBox18.Text = nilaipertama1.ToString("N2")
        MetroTextBox19.Text = kuantitikedua1
        MetroTextBox20.Text = nilaikedua1.ToString("N2")
        MetroTextBox21.Text = kuantitiketiga1
        MetroTextBox22.Text = nilaiketiga1.ToString("N2")
        MetroTextBox23.Text = kuantitikeempat1
        MetroTextBox24.Text = nilaikeempat1.ToString("N2")
    End Sub

    'Public Sub keluaran()
    '    Dim oleConn As System.Data.OleDb.OleDbConnection
    '    Dim da As OleDb.OleDbDataAdapter
    '    ds = New DataSet
    '    oleConn = New System.Data.OleDb.OleDbConnection
    '    oleConn.ConnectionString = My.Settings.manageConnectionString
    '    Try
    '        oleConn.Open()
    '        da = New OleDb.OleDbDataAdapter("select kod,nama,harga,kuantiti,jumlah,tarikh,suku from kurang where kod like '%" & MetroTextBox42.Text & "%'", oleConn)
    '        da.Fill(ds, "keluar")
    '        oleConn.Close()
    '    Catch ex As Exception
    '        MsgBox("Tiada Data Dalam Database.")
    '    End Try
    '    'Dim kuantitipertama As Decimal = 0
    '    'Dim nilaipertama As Decimal = 0
    '    'Dim kuantitikedua As Decimal = 0
    '    'Dim nilaikedua As Decimal = 0
    '    'Dim kuantitiketiga As Decimal = 0
    '    'Dim nilaiketiga As Decimal = 0
    '    'Dim kuantitikeempat As Decimal = 0
    '    'Dim nilaikeempat As Decimal = 0
    '    For row As Integer = 0 To ds.Tables("keluar").Rows.Count - 1
    '        If ds.Tables("keluar").Rows(row).Item("suku") = 1 Then

    '            kuantitipertama1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
    '            nilaipertama1 += ds.Tables("keluar").Rows(row).Item("jumlah")

    '        ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 2 Then

    '            kuantitikedua1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
    '            nilaikedua1 += ds.Tables("keluar").Rows(row).Item("jumlah")

    '        ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 3 Then

    '            kuantitiketiga1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
    '            nilaiketiga1 += ds.Tables("keluar").Rows(row).Item("jumlah")

    '        ElseIf ds.Tables("keluar").Rows(row).Item("suku") = 4 Then

    '            kuantitikeempat1 += ds.Tables("keluar").Rows(row).Item("kuantiti")
    '            nilaikeempat1 += ds.Tables("keluar").Rows(row).Item("jumlah")

    '        End If
    '    Next
    '    MetroTextBox17.Text = kuantitipertama1
    '    MetroTextBox18.Text = nilaipertama1.ToString("N2")
    '    MetroTextBox19.Text = kuantitikedua1
    '    MetroTextBox20.Text = nilaikedua1.ToString("N2")
    '    MetroTextBox21.Text = kuantitiketiga1
    '    MetroTextBox22.Text = nilaiketiga1.ToString("N2")
    '    MetroTextBox23.Text = kuantitikeempat1
    '    MetroTextBox24.Text = nilaikeempat1.ToString("N2")
    'End Sub

    Public Sub stoksemasa(ByVal i As Decimal, ByVal ii As Decimal, ByVal iii As Decimal)
        jumlah = (i + ii) - iii
        sedia = jumlah
    End Sub

    Public Sub kira()
        MetroTextBox1.Text = sedia
        stoksemasa(sedia, kuantitipertama, kuantitipertama1)
        MetroTextBox25.Text = jumlah
        MetroTextBox3.Text = sedia
        stoksemasa(sedia, kuantitikedua, kuantitikedua1)
        MetroTextBox27.Text = jumlah
        MetroTextBox5.Text = sedia
        stoksemasa(sedia, kuantitiketiga, kuantitiketiga1)
        MetroTextBox29.Text = jumlah
        MetroTextBox7.Text = sedia
        stoksemasa(sedia, kuantitikeempat, kuantitikeempat1)
        MetroTextBox31.Text = jumlah

        MetroTextBox2.Text = sedia1
        stoksemasa(sedia1, nilaipertama, nilaipertama1)
        MetroTextBox26.Text = jumlah
        kadar = nilaipertama1 / ((sedia1 + jumlah) / 2)
        MetroTextBox33.Text = kadar.ToString("N2")

        MetroTextBox4.Text = sedia
        stoksemasa(sedia, nilaikedua, nilaikedua1)
        MetroTextBox28.Text = jumlah
        kadar1 = nilaikedua1 / ((sedia + jumlah) / 2)
        MetroTextBox34.Text = kadar1.ToString("N2")

        MetroTextBox6.Text = sedia
        stoksemasa(sedia, nilaiketiga, nilaiketiga1)
        MetroTextBox30.Text = jumlah
        kadar2 = nilaiketiga1 / ((sedia + jumlah) / 2)
        MetroTextBox35.Text = kadar2.ToString("N2")

        MetroTextBox8.Text = sedia
        stoksemasa(sedia, nilaikeempat, nilaikeempat1)
        MetroTextBox32.Text = jumlah
        kadar3 = nilaikeempat1 / ((sedia + jumlah) / 2)
        MetroTextBox36.Text = kadar3.ToString("N2")

        jumlahkadar = kadar + kadar1 + kadar2 + kadar3
        MetroTextBox41.Text = jumlahkadar.ToString("N2")

        totalkuantiti = kuantitipertama + kuantitikedua + kuantitiketiga + kuantitikeempat
        totalnilai = nilaipertama + nilaikedua + nilaiketiga + nilaikeempat
        totalkuantiti1 = kuantitipertama1 + kuantitikedua1 + kuantitiketiga1 + kuantitikeempat1
        totalnilai1 = nilaipertama1 + nilaikedua1 + nilaiketiga1 + nilaikeempat1
        MetroTextBox37.Text = totalkuantiti
        MetroTextBox38.Text = totalnilai.ToString("N2")
        MetroTextBox39.Text = totalkuantiti1
        MetroTextBox40.Text = totalnilai1.ToString("N2")
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        terimaan()
        keluaran()
        kira()
    End Sub

    Private Sub MetroButton13_Click(sender As Object, e As EventArgs) Handles MetroButton13.Click
        Me.Close()
    End Sub

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        Dim acsconn As System.Data.OleDb.OleDbConnection
        acsconn = New System.Data.OleDb.OleDbConnection
        acsconn.ConnectionString = My.Settings.manageConnectionString
        Dim cmd As New OleDb.OleDbCommand
        Try
            acsconn.Open()
            cmd.Connection = acsconn
            cmd.CommandText = "insert into stok(kod,sa1,sa11,sa2,sa22,sa3,sa33,sa4,sa44,b1,b11,b2,b22,b3,b33,b4,b44,k1,k11,k2,k22,k3,k33,k4,k44,ss1,ss11,ss2,ss22,ss3,ss33,ss4,ss44,kps1,kps2,kps3,kps4,ntb1,ntb11,ntk1,ntk11,kps) values(@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n,@o,@p,@q,@r,@s,@t,@u,@v,@w,@x,@y,@z,@aa,@bb,@cc,@dd,@ee,@ff,@gg,@hh,@ii,@jj,@kk,@ll,@mm,@nn,@oo,@pp);"
            cmd.Parameters.AddWithValue("@a", MetroTextBox42.Text)
            cmd.Parameters.AddWithValue("@b", MetroTextBox1.Text)
            cmd.Parameters.AddWithValue("@c", MetroTextBox2.Text)
            cmd.Parameters.AddWithValue("@d", MetroTextBox3.Text)
            cmd.Parameters.AddWithValue("@e", MetroTextBox4.Text)
            cmd.Parameters.AddWithValue("@f", MetroTextBox5.Text)
            cmd.Parameters.AddWithValue("@g", MetroTextBox6.Text)
            cmd.Parameters.AddWithValue("@h", MetroTextBox7.Text)
            cmd.Parameters.AddWithValue("@i", MetroTextBox8.Text)
            cmd.Parameters.AddWithValue("@j", MetroTextBox9.Text)
            cmd.Parameters.AddWithValue("@k", MetroTextBox10.Text)
            cmd.Parameters.AddWithValue("@l", MetroTextBox11.Text)
            cmd.Parameters.AddWithValue("@m", MetroTextBox12.Text)
            cmd.Parameters.AddWithValue("@n", MetroTextBox13.Text)
            cmd.Parameters.AddWithValue("@o", MetroTextBox14.Text)
            cmd.Parameters.AddWithValue("@p", MetroTextBox15.Text)
            cmd.Parameters.AddWithValue("@q", MetroTextBox16.Text)
            cmd.Parameters.AddWithValue("@r", MetroTextBox17.Text)
            cmd.Parameters.AddWithValue("@s", MetroTextBox18.Text)
            cmd.Parameters.AddWithValue("@t", MetroTextBox19.Text)
            cmd.Parameters.AddWithValue("@u", MetroTextBox20.Text)
            cmd.Parameters.AddWithValue("@v", MetroTextBox21.Text)
            cmd.Parameters.AddWithValue("@w", MetroTextBox22.Text)
            cmd.Parameters.AddWithValue("@x", MetroTextBox23.Text)
            cmd.Parameters.AddWithValue("@y", MetroTextBox24.Text)
            cmd.Parameters.AddWithValue("@z", MetroTextBox25.Text)
            cmd.Parameters.AddWithValue("@aa", MetroTextBox26.Text)
            cmd.Parameters.AddWithValue("@bb", MetroTextBox27.Text)
            cmd.Parameters.AddWithValue("@cc", MetroTextBox28.Text)
            cmd.Parameters.AddWithValue("@dd", MetroTextBox29.Text)
            cmd.Parameters.AddWithValue("@ee", MetroTextBox30.Text)
            cmd.Parameters.AddWithValue("@ff", MetroTextBox31.Text)
            cmd.Parameters.AddWithValue("@gg", MetroTextBox32.Text)
            cmd.Parameters.AddWithValue("@hh", MetroTextBox33.Text)
            cmd.Parameters.AddWithValue("@ii", MetroTextBox34.Text)
            cmd.Parameters.AddWithValue("@jj", MetroTextBox35.Text)
            cmd.Parameters.AddWithValue("@kk", MetroTextBox36.Text)
            cmd.Parameters.AddWithValue("@ll", MetroTextBox37.Text)
            cmd.Parameters.AddWithValue("@mm", MetroTextBox38.Text)
            cmd.Parameters.AddWithValue("@nn", MetroTextBox39.Text)
            cmd.Parameters.AddWithValue("@oo", MetroTextBox40.Text)
            cmd.Parameters.AddWithValue("@pp", MetroTextBox41.Text)
            cmd.ExecuteNonQuery()
            acsconn.Close()
            MsgBox("berjaya simpan")
            Form1.kedudukan(Form1.ListView5)
            Me.Close()
        Catch ex As Exception
            MsgBox("xberjaya. sila cuba lg.")
        End Try
    End Sub

    'Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    '    thisdate = Date.Now.ToString
    '    thisyear = Year(thisdate)
    '    MsgBox(thisyear.ToString)
    'End Sub
End Class