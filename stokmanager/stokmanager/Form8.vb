Public Class Form8
    Public Shared ds As DataSet
    Public Shared itemout As DataTable
    Dim row As Integer
    Dim row1 As Integer
    Dim a As Decimal = 0
    Dim b As Decimal = 0
    Dim c As Decimal = 0
    Dim jumlah As Decimal = 0
    Dim baki As Decimal = 0
    Dim harga As String
    Dim namapembekal As String
    Dim alamatpembekal As String
    Dim poskod As String
    Dim negeri As String
    Dim nopesanan As String
    Dim kuantitipesanan As String
    Dim kuantititerima As String
    Dim suku As String
    Private Sub Form8_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim thisdate As Date
        Dim thismonth As Integer
        thisdate = Date.Now.ToString
        thismonth = Month(thisdate)
        If thismonth = "1" Or thismonth = "2" Or thismonth = "3" Then
            suku = "1"
        ElseIf thismonth = "4" Or thismonth = "5" Or thismonth = "6" Then
            suku = "2"
        ElseIf thismonth = "7" Or thismonth = "8" Or thismonth = "9" Then
            suku = "3"
        ElseIf thismonth = "10" Or thismonth = "11" Or thismonth = "12" Then
            suku = "4"
        End If
        table1()
        MetroTextBox4.Text = Form1.Label2.Text
        ShowDataInLvw1(itemout, ListView1)
    End Sub
    Public Sub table1()
        itemout = New DataTable("OUT")

        Dim column1 As DataColumn = New DataColumn("BARKOD")
        column1.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column1)

        Dim column2 As DataColumn = New DataColumn("NAMA ITEM")
        column2.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column2)

        Dim column14 As DataColumn = New DataColumn("HARGA")
        column14.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column14)

        Dim column3 As DataColumn = New DataColumn("KUANTITI DIPESAN")
        column3.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column3)

        Dim column4 As DataColumn = New DataColumn("KUANTITI DILULUSKAN")
        column4.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column4)

        Dim column5 As DataColumn = New DataColumn("BAKI KUANTITI")
        column5.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column5)

        Dim column8 As DataColumn = New DataColumn("JUMLAH")
        column8.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column8)

        Dim column6 As DataColumn = New DataColumn("CATATAN")
        column6.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column6)

        Dim column7 As DataColumn = New DataColumn("TARIKH")
        column7.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column7)

        Dim column9 As DataColumn = New DataColumn("NAMA PEMBEKAL")
        column9.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column9)

        Dim column10 As DataColumn = New DataColumn("ALAMAT PEMBEKAL")
        column10.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column10)

        Dim column11 As DataColumn = New DataColumn("POSKOD")
        column11.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column11)

        Dim column12 As DataColumn = New DataColumn("NEGERI")
        column12.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column12)

        Dim column13 As DataColumn = New DataColumn("NO.PESANAN")
        column13.DataType = System.Type.GetType("System.String")
        itemout.Columns.Add(column13)

        'Dim column14 As DataColumn = New DataColumn("KUANTITI PESANAN")
        'column14.DataType = System.Type.GetType("System.String")
        'itemout.Columns.Add(column14)

        'Dim column15 As DataColumn = New DataColumn("KUANTITI TERIMA")
        'column15.DataType = System.Type.GetType("System.String")
        'itemout.Columns.Add(column15)
    End Sub

    Private Sub MetroButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton8.Click
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Try
            oleConn.Open()
            da = New OleDb.OleDbDataAdapter("select * from detail where kod like '%" & MetroTextBox12.Text & "%'", oleConn)
            da.Fill(ds, "infox")
            oleConn.Close()
            MetroTextBox1.Text = ds.Tables("infox").Rows(row)("nama")
            MetroTextBox3.Text = ds.Tables("infox").Rows(row)("harga")
            MetroTextBox2.Text = ds.Tables("infox").Rows(row)("kuantiti")
            a = ds.Tables("infox").Rows(row)("kuantiti")
            harga = ds.Tables("infox").Rows(row)("harga")
            namapembekal = ds.Tables("infox").Rows(row)("namapembekal")
            alamatpembekal = ds.Tables("infox").Rows(row)("alamatpembekal")
            poskod = ds.Tables("infox").Rows(row)("poskod")
            negeri = ds.Tables("infox").Rows(row)("negeri")
            nopesanan = ds.Tables("infox").Rows(row)("pesanan")
            ' kuantitipesanan = ds.Tables("infox").Rows(row)("kuantitipesan")
            'kuantititerima = ds.Tables("infox").Rows(row)("kuantititerima")
            ' MetroTextBox8.Text = ds.Tables("infox").Rows(row)("tarikh")
           
        Catch ex As Exception
            MsgBox("Item Tiada Dalam Simpanan. Sila Masukkan Item Mengikut Seksyen Item Masuk Dahulu.")
            MetroTextBox1.Text = ""
            MetroTextBox2.Text = ""
            MetroTextBox3.Text = ""
            ' MetroTextBox8.Text = ""
            ' MetroLabel16.Visible = False
            'MetroTextBox15.Visible = False
            MetroTextBox12.Text = ""
            MetroTextBox12.Focus()
            'dahade = False
        End Try
    End Sub
    Public Sub ShowDataInLvw1(ByVal data As DataTable, ByVal lvw As ListView)
        lvw.View = View.Details
        lvw.GridLines = True
        lvw.Columns.Clear()
        lvw.Items.Clear()
        Dim x As Integer = 0
        For Each col As DataColumn In data.Columns
            x = x + 1
            Select Case x

                Case 1
                    lvw.Columns.Add(col.ToString, 100, HorizontalAlignment.Center)
                Case 2
                    lvw.Columns.Add(col.ToString, 150, HorizontalAlignment.Center)
                Case 3
                    lvw.Columns.Add(col.ToString, 150, HorizontalAlignment.Center)
                Case 4
                    lvw.Columns.Add(col.ToString, 150, HorizontalAlignment.Center)
                Case 5
                    lvw.Columns.Add(col.ToString, 150, HorizontalAlignment.Center)
                Case 6
                    lvw.Columns.Add(col.ToString, 150, HorizontalAlignment.Center)
                Case 7
                    lvw.Columns.Add(col.ToString, 240, HorizontalAlignment.Center)
                Case 8
                    lvw.Columns.Add(col.ToString, 150, HorizontalAlignment.Center)
                Case 9
                    lvw.Columns.Add(col.ToString, 140, HorizontalAlignment.Center)
                Case 10
                    lvw.Columns.Add(col.ToString, 140, HorizontalAlignment.Center)
                Case 11
                    lvw.Columns.Add(col.ToString, 140, HorizontalAlignment.Center)
                Case 12
                    lvw.Columns.Add(col.ToString, 140, HorizontalAlignment.Center)
                Case 13
                    lvw.Columns.Add(col.ToString, 260, HorizontalAlignment.Center)
                Case Else
                    lvw.Columns.Add(col.ToString, 100, HorizontalAlignment.Center)
            End Select
        Next

        For Each row As DataRow In data.Rows
            Dim lst As ListViewItem
            lst = lvw.Items.Add(row(0))
            For i As Integer = 1 To data.Columns.Count - 1
                lst.SubItems.Add(row(i))
            Next
        Next
    End Sub

    Private Sub MetroButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton1.Click
        b = Convert.ToDecimal(MetroTextBox6.Text)
        If b > a Then
            MsgBox("Kuantiti Pesanan Lebih Daripada Stok Yang Ada.")
            MetroTextBox5.Text = ""
        Else
            baki = a - b
            ' MsgBox(baki.ToString)
            c = Convert.ToDecimal(MetroTextBox3.Text)
            jumlah = baki * c
            'MsgBox(jumlah.ToString)
            Dim Rowx As DataRow
            Rowx = itemout.NewRow()
            Rowx.Item("BARKOD") = MetroTextBox12.Text
            Rowx.Item("NAMA ITEM") = MetroTextBox1.Text
            Rowx.Item("HARGA") = ds.Tables("infox").Rows(row)("harga")
            Rowx.Item("KUANTITI DIPESAN") = MetroTextBox5.Text
            Rowx.Item("KUANTITI DILULUSKAN") = MetroTextBox6.Text
            Rowx.Item("BAKI KUANTITI") = baki.ToString
            Rowx.Item("JUMLAH") = jumlah.ToString("N2")
            Rowx.Item("CATATAN") = MetroTextBox8.Text
            Rowx.Item("TARIKH") = MetroTextBox4.Text
            Rowx.Item("NAMA PEMBEKAL") = namapembekal
            Rowx.Item("ALAMAT PEMBEKAL") = alamatpembekal
            Rowx.Item("POSKOD") = poskod
            Rowx.Item("NEGERI") = negeri
            Rowx.Item("NO.PESANAN") = nopesanan
            'Rowx.Item("KUANTITI PESANAN") = kuantitipesanan
            'Rowx.Item("KUANTITI TERIMA") = kuantititerima
            itemout.Rows.Add(Rowx)
            ShowDataInLvw1(itemout, ListView1)
            MetroTextBox1.Text = ""
            MetroTextBox2.Text = ""
            MetroTextBox3.Text = ""
            MetroTextBox4.Text = ""
            MetroTextBox12.Text = ""
            MetroTextBox5.Text = ""
            MetroTextBox6.Text = ""
            MetroTextBox8.Text = ""
        End If
    End Sub

    Private Sub MetroButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton4.Click
        insert()
        adddata()
        printout()
        Form1.readdatashow(Form1.ListView1)
        Form1.readdatashow(Form1.ListView2)
        Form1.readdatashow(Form1.ListView3)
        MetroTextBox12.Text = ""
        MetroTextBox1.Text = ""
        MetroTextBox2.Text = ""
        MetroTextBox3.Text = ""
        'MetroTextBox4.Text = ""
        MetroTextBox5.Text = ""
        MetroTextBox6.Text = ""
        MetroTextBox8.Text = ""
        Me.Close()
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        row1 = ListView1.SelectedItems(0).Index
        itemout.Rows(row1).Delete()
        ShowDataInLvw1(itemout, ListView1)
    End Sub
    Public Sub insert()
        For row As Integer = 0 To itemout.Rows.Count - 1
            Dim acsconn As System.Data.OleDb.OleDbConnection
            acsconn = New System.Data.OleDb.OleDbConnection
            acsconn.ConnectionString = My.Settings.manageConnectionString
            Dim cmd As New OleDb.OleDbCommand
            Try
                acsconn.Open()
                cmd.Connection = acsconn
                cmd.CommandText = "insert into kurang(kod,nama,harga,kuantiti,tarikh,jumlah,namapembekal,alamatpembekal,poskod,negeri,pesanan,kuantitipesan,kuantititerima) values(@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m);"
                cmd.Parameters.AddWithValue("@a", itemout.Rows(row).Item("BARKOD"))
                cmd.Parameters.AddWithValue("@b", itemout.Rows(row).Item("NAMA ITEM"))
                cmd.Parameters.AddWithValue("@c", itemout.Rows(row).Item("HARGA"))
                cmd.Parameters.AddWithValue("@d", itemout.Rows(row).Item("BAKI KUANTITI"))
                cmd.Parameters.AddWithValue("@e", itemout.Rows(row).Item("TARIKH"))
                cmd.Parameters.AddWithValue("@f", itemout.Rows(row).Item("JUMLAH"))
                cmd.Parameters.AddWithValue("@g", itemout.Rows(row).Item("NAMA PEMBEKAL"))
                cmd.Parameters.AddWithValue("@h", itemout.Rows(row).Item("ALAMAT PEMBEKAL"))
                cmd.Parameters.AddWithValue("@i", itemout.Rows(row).Item("POSKOD"))
                cmd.Parameters.AddWithValue("@j", itemout.Rows(row).Item("NEGERI"))
                cmd.Parameters.AddWithValue("@k", itemout.Rows(row).Item("NO.PESANAN"))
                cmd.Parameters.AddWithValue("@l", itemout.Rows(row).Item("KUANTITI DIPESAN"))
                cmd.Parameters.AddWithValue("@m", itemout.Rows(row).Item("KUANTITI DILULUSKAN"))
                'cmd.Parameters.AddWithValue("@g", id)
                cmd.ExecuteNonQuery()
                acsconn.Close()
                MsgBox("Pendaftaran Item Berjaya. Data Telah Disimpan.")
                Form1.readdatashow(Form1.ListView2)
          
                Me.Close()
            Catch ex As Exception
                MsgBox(ErrorToString)
            End Try
        Next
    End Sub
    Public Sub adddata()
        For row As Integer = 0 To itemout.Rows.Count - 1

            Dim acsconn As System.Data.OleDb.OleDbConnection
            acsconn = New System.Data.OleDb.OleDbConnection
            acsconn.ConnectionString = My.Settings.manageConnectionString
            Dim cmd As New OleDb.OleDbCommand
            Try
                acsconn.Open()
                cmd.Connection = acsconn
                cmd.CommandText = "update detail set kod=@a,nama=@b,harga=@c,kuantiti=@d,jumlah=@e,tarikh=@f,namapembekal=@g,alamatpembekal=@h,poskod=@i,pesanan=@k,negeri=@j,suku=@l where kod = @a;"
                cmd.Parameters.AddWithValue("@a", itemout.Rows(row).Item("BARKOD"))
                cmd.Parameters.AddWithValue("@b", itemout.Rows(row).Item("NAMA ITEM"))
                cmd.Parameters.AddWithValue("@c", itemout.Rows(row).Item("HARGA"))
                cmd.Parameters.AddWithValue("@d", itemout.Rows(row).Item("BAKI KUANTITI"))
                cmd.Parameters.AddWithValue("@e", itemout.Rows(row).Item("JUMLAH"))
                cmd.Parameters.AddWithValue("@f", itemout.Rows(row).Item("TARIKH"))
                cmd.Parameters.AddWithValue("@g", itemout.Rows(row).Item("NAMA PEMBEKAL"))
                cmd.Parameters.AddWithValue("@h", itemout.Rows(row).Item("ALAMAT PEMBEKAL"))
                cmd.Parameters.AddWithValue("@i", itemout.Rows(row).Item("POSKOD"))
                cmd.Parameters.AddWithValue("@j", itemout.Rows(row).Item("NEGERI"))
                cmd.Parameters.AddWithValue("@k", itemout.Rows(row).Item("NO.PESANAN"))
                cmd.Parameters.AddWithValue("@l", suku)
                cmd.ExecuteNonQuery()
                acsconn.Close()

            
            Catch ex As Exception
                MsgBox("gagal ubah")
            End Try
        Next
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        MetroTextBox12.Text = ""
        MetroTextBox1.Text = ""
        MetroTextBox2.Text = ""
        MetroTextBox3.Text = ""
        MetroTextBox4.Text = ""
        MetroTextBox5.Text = ""
        MetroTextBox6.Text = ""
        MetroTextBox8.Text = ""
    End Sub
    Public Sub printout()
        PrintPreviewDialog1.Document = PrintDocument1 'PrintPreviewDialog associate with PrintDocument.

        DirectCast(DirectCast(PrintPreviewDialog1.Controls(1), ToolStrip).Items(0), ToolStripButton).Enabled = False
        PrintPreviewDialog1.ShowDialog()

        PrintDialog1.Document = PrintDocument1 'PrintDialog associate with PrintDocument.

        If PrintDialog1.ShowDialog() = DialogResult.OK Then

            PrintDocument1.Print()

        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim drawFont As New Font("Arial", 12)
        Dim drawFont1 As New Font("Arial", 10)
        Dim drawFont2 As New Font("Arial", 14)
        Dim drawBrush As New SolidBrush(Color.DarkBlue)
        Dim drawbrush1 As New SolidBrush(Color.Black)
        Dim drawFormat As New StringFormat
        Dim blackPen As New Pen(Color.Black, 3)
        e.Graphics.DrawString("KEW.PS-11", drawFont, drawBrush, 700.0F, 15.0F, drawFormat)
        e.Graphics.DrawString("BORANG PENERIMAAN BARANG (BTB)", drawFont, drawBrush, 230.0F, 25.0F, drawFormat)
        e.Graphics.DrawString("(Tatacara Pengurusan Stor 43)", drawFont1, drawBrush, 300.0F, 45.0F, drawFormat)

       
        'horizon line
        e.Graphics.DrawLine(blackPen, 50.0F, 130.0F, 790.0F, 130.0F)
        e.Graphics.DrawLine(blackPen, 410.0F, 155.0F, 635.0F, 155.0F)
        e.Graphics.DrawLine(blackPen, 50.0F, 205.0F, 790.0F, 205.0F)
        
        'e.Graphics.DrawLine(blackPen, 120.0F, 250.0F, 750.0F, 250.0F)
        'e.Graphics.DrawLine(blackPen, 120.0F, 280.0F, 750.0F, 280.0F)

        'vertical line
        e.Graphics.DrawLine(blackPen, 50.0F, 130.0F, 50.0F, 205.0F)
        e.Graphics.DrawLine(blackPen, 300.0F, 130.0F, 300.0F, 205.0F)
        e.Graphics.DrawLine(blackPen, 410.0F, 130.0F, 410.0F, 205.0F)
        e.Graphics.DrawLine(blackPen, 525.0F, 155.0F, 525.0F, 205.0F)
        e.Graphics.DrawLine(blackPen, 635.0F, 130.0F, 635.0F, 205.0F)
        e.Graphics.DrawLine(blackPen, 790.0F, 130.0F, 790.0F, 205.0F)

        Dim i As Integer = 0
        Dim height As Integer = 205.0F
        For row As Integer = 0 To itemout.Rows.Count - 1
            e.Graphics.DrawLine(blackPen, 50.0F, 245.0F + i, 790.0F, 245.0F + i)
            'vertical
            e.Graphics.DrawLine(blackPen, 50.0F, 205.0F, 50.0F, 245.0F + i)
            e.Graphics.DrawLine(blackPen, 300.0F, 205.0F, 300.0F, 245.0F + i)
            e.Graphics.DrawLine(blackPen, 410.0F, 205.0F, 410.0F, 245.0F + i)
            e.Graphics.DrawLine(blackPen, 525.0F, 205.0F, 525.0F, 245.0F + i)
            e.Graphics.DrawLine(blackPen, 635.0F, 205.0F, 635.0F, 245.0F + i)
            e.Graphics.DrawLine(blackPen, 790.0F, 205.0F, 790.0F, 245.0F + i)

            'detail
            e.Graphics.DrawString(itemout.Rows(row).Item("NAMA ITEM") + " (" + itemout.Rows(row).Item("BARKOD") + ")", drawFont, drawbrush1, 60.0F, 215.0F + i, drawFormat)
            e.Graphics.DrawString(itemout.Rows(row).Item("KUANTITI DIPESAN"), drawFont, drawbrush1, 305.0F, 215.0F + i, drawFormat)
            e.Graphics.DrawString(itemout.Rows(row).Item("KUANTITI DILULUSKAN"), drawFont, drawbrush1, 415.0F, 215.0F + i, drawFormat)
            e.Graphics.DrawString(itemout.Rows(row).Item("BAKI KUANTITI"), drawFont, drawbrush1, 530.0F, 215.0F + i, drawFormat)
            e.Graphics.DrawString(itemout.Rows(row).Item("CATATAN"), drawFont, drawbrush1, 640.0F, 215.0F + i, drawFormat)
            i = i + 40
            height = height + 40
        Next

        'e.Graphics.DrawLine(blackPen, 50.0F, 280.0F, 790.0F, 280.0F)
        e.Graphics.DrawLine(blackPen, 50.0F, height + 255.0F, 790.0F, height + 255.0F)

        'vertical
        e.Graphics.DrawLine(blackPen, 50.0F, height, 50.0F, height + 255.0F)
        e.Graphics.DrawLine(blackPen, 410.0F, height, 410.0F, height + 255.0F)
        e.Graphics.DrawLine(blackPen, 790.0F, height, 790.0F, height + 255.0F)

        'penerangan
        e.Graphics.DrawString("Pegawai Pelulus", drawFont, drawBrush, 450.0F, 133.0F, drawFormat)
        e.Graphics.DrawString("Pengeluaran Stok", drawFont2, drawBrush, 60.0F, 145.0F, drawFormat)
        e.Graphics.DrawString("Kuantiti", drawFont2, drawBrush, 310.0F, 145.0F, drawFormat)
        e.Graphics.DrawString("  Kuantiti " + Environment.NewLine + "Diluluskan", drawFont, drawBrush, 420.0F, 160.0F, drawFormat)
        e.Graphics.DrawString("  Baki " + Environment.NewLine + "Kuantiti", drawFont, drawBrush, 540.0F, 160.0F, drawFormat)
        e.Graphics.DrawString("Catatan", drawFont2, drawBrush, 660.0F, 145.0F, drawFormat)

        e.Graphics.DrawString("..........................", drawFont2, drawBrush, 60.0F, height + 145.0F, drawFormat)
        e.Graphics.DrawString("(Tandatangan Pemohon)", drawFont2, drawBrush, 60.0F, height + 165.0F, drawFormat)
        e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 60.0F, height + 185.0F, drawFormat)
        e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 60.0F, height + 205.0F, drawFormat)
        e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 60.0F, height + 225.0F, drawFormat)

        e.Graphics.DrawString("..........................", drawFont2, drawBrush, 410.0F, height + 145.0F, drawFormat)
        e.Graphics.DrawString("(Tandatangan Pegawai Pelulus)", drawFont2, drawBrush, 410.0F, height + 165.0F, drawFormat)
        e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 410.0F, height + 185.0F, drawFormat)
        e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 410.0F, height + 205.0F, drawFormat)
        e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 410.0F, height + 225.0F, drawFormat)



       
        '-------------------------------------------------------------------------------------------------------------------------
        'penerangan
        e.Graphics.DrawString("Kemaskini Rekod:", drawFont2, drawBrush, 55.0F, height + 275.0F, drawFormat)
        e.Graphics.DrawString("Stok telah dikeluarkan dan" + Environment.NewLine + "direkod di Kad Petak No............", drawFont, drawBrush, 55.0F, height + 295.0F, drawFormat)
        e.Graphics.DrawString(".......................", drawFont2, drawBrush, 55.0F, height + 410.0F, drawFormat)
        e.Graphics.DrawString("(Tandatangan Pegawai Stor)", drawFont2, drawBrush, 55.0F, height + 430.0F, drawFormat)
        e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 55.0F, height + 450.0F, drawFormat)
        e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 55.0F, height + 470.0F, drawFormat)
        e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 55.0F, height + 490.0F, drawFormat)

        e.Graphics.DrawString("Perakuan Penerimaan:", drawFont2, drawBrush, 415.0F, height + 275.0F, drawFormat)
        e.Graphics.DrawString("Disahkan bahawa stok yang diluluskan" + Environment.NewLine + "telah diterima.", drawFont, drawBrush, 415.0F, height + 295.0F, drawFormat)
        e.Graphics.DrawString(".......................", drawFont2, drawBrush, 415.0F, height + 410.0F, drawFormat)
        e.Graphics.DrawString("(Tandatangan Pemohon)", drawFont2, drawBrush, 415.0F, height + 430.0F, drawFormat)
        e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 415.0F, height + 450.0F, drawFormat)
        e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 415.0F, height + 470.0F, drawFormat)
        e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 415.0F, height + 490.0F, drawFormat)

        'horizontal line
        e.Graphics.DrawLine(blackPen, 50.0F, height + 270.0F, 790.0F, height + 270.0F)
        e.Graphics.DrawLine(blackPen, 50.0F, height + 515.0F, 790.0F, height + 515.0F)

        'vertical line
        e.Graphics.DrawLine(blackPen, 50.0F, height + 270.0F, 50.0F, height + 515.0F)
        e.Graphics.DrawLine(blackPen, 410.0F, height + 270.0F, 410.0F, height + 515.0F)
        e.Graphics.DrawLine(blackPen, 790.0F, height + 270.0F, 790.0F, height + 515.0F)
    End Sub

   
    Private Sub MetroButton5_Click(sender As Object, e As EventArgs)
        printout()
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        Me.Close()
    End Sub
End Class