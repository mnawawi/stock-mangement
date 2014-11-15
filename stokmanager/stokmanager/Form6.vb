Public Class Form6

    Public Shared ds As DataSet
    Public Shared datall As DataTable
    Dim a As Decimal = 0
    Dim b As Decimal = 0
    Dim c As Decimal = 0
    Dim d As Decimal = 0
    Dim jumlah As Decimal = 0
    Dim jumlah1 As Decimal = 0
    Dim kuantiti As Decimal = 0
    Dim kuantiti1 As Decimal = 0
    Dim row1 As Integer
    Dim dahade As Boolean = False
    Dim suku As String

    Public Sub table1()
        datall = New DataTable("ITEM")

        Dim column1 As DataColumn = New DataColumn("BARKOD")
        column1.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column1)

        Dim column2 As DataColumn = New DataColumn("NAMA ITEM")
        column2.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column2)

        Dim column3 As DataColumn = New DataColumn("HARGA")
        column3.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column3)

        Dim column4 As DataColumn = New DataColumn("KUANTITI")
        column4.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column4)

        Dim column13 As DataColumn = New DataColumn("JUMLAH")
        column13.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column13)

        Dim column14 As DataColumn = New DataColumn("TARIKH")
        column14.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column14)

        Dim column5 As DataColumn = New DataColumn("NAMA PEMBEKAL")
        column5.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column5)

        Dim column6 As DataColumn = New DataColumn("ALAMAT PEMBEKAL")
        column6.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column6)

        Dim column7 As DataColumn = New DataColumn("POSKOD")
        column7.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column7)

        Dim column8 As DataColumn = New DataColumn("NEGERI")
        column8.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column8)

        Dim column9 As DataColumn = New DataColumn("NO.PESANAN")
        column9.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column9)

        Dim column10 As DataColumn = New DataColumn("KUANTITI PESANAN")
        column10.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column10)

        Dim column11 As DataColumn = New DataColumn("KUANTITI TERIMA")
        column11.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column11)

        Dim column12 As DataColumn = New DataColumn("CATATAN")
        column12.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column12)

        Dim column16 As DataColumn = New DataColumn("KUANTITI1")
        column16.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column16)

        Dim column17 As DataColumn = New DataColumn("JUMLAH1")
        column17.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column17)
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

    Private Sub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '  readdatashow(ListView1)
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
        MetroTextBox7.Text = Form1.Label2.Text
        table1()
        ShowDataInLvw1(datall, ListView1)
    End Sub

    Private Sub MetroButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton1.Click
        Dim Rowx As DataRow
        If dahade Then

            Rowx = datall.NewRow()
            Rowx.Item("BARKOD") = MetroTextBox3.Text
            Rowx.Item("NAMA ITEM") = MetroTextBox4.Text
            Rowx.Item("HARGA") = MetroTextBox5.Text
            Rowx.Item("KUANTITI") = kuantiti
            Rowx.Item("JUMLAH") = MetroTextBox12.Text
            Rowx.Item("TARIKH") = MetroTextBox7.Text
            Rowx.Item("NAMA PEMBEKAL") = MetroTextBox13.Text
            Rowx.Item("ALAMAT PEMBEKAL") = MetroTextBox10.Text
            Rowx.Item("POSKOD") = MetroTextBox9.Text
            Rowx.Item("NEGERI") = ComboBox1.Text
            Rowx.Item("NO.PESANAN") = MetroTextBox8.Text
            Rowx.Item("KUANTITI PESANAN") = MetroTextBox1.Text
            Rowx.Item("KUANTITI TERIMA") = MetroTextBox11.Text
            Rowx.Item("CATATAN") = MetroTextBox2.Text
            Rowx.Item("KUANTITI1") = kuantiti1
            Rowx.Item("JUMLAH1") = jumlah1
            datall.Rows.Add(Rowx)
            ShowDataInLvw1(datall, ListView1)
        Else
            Dim acsconn As System.Data.OleDb.OleDbConnection
            acsconn = New System.Data.OleDb.OleDbConnection
            acsconn.ConnectionString = My.Settings.manageConnectionString
            Dim cmd As New OleDb.OleDbCommand
            Try
                acsconn.Open()
                cmd.Connection = acsconn
                cmd.CommandText = "insert into detail(kod,nama,harga,kuantiti,tarikh,jumlah,namapembekal,alamatpembekal,poskod,negeri,pesanan,suku) values(@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l);"
                cmd.Parameters.AddWithValue("@a", MetroTextBox3.Text)
                cmd.Parameters.AddWithValue("@b", MetroTextBox4.Text)
                cmd.Parameters.AddWithValue("@c", MetroTextBox5.Text)
                cmd.Parameters.AddWithValue("@d", MetroTextBox6.Text)
                cmd.Parameters.AddWithValue("@e", MetroTextBox7.Text)
                cmd.Parameters.AddWithValue("@f", MetroTextBox12.Text)
                cmd.Parameters.AddWithValue("@g", MetroTextBox13.Text)
                cmd.Parameters.AddWithValue("@h", MetroTextBox10.Text)
                cmd.Parameters.AddWithValue("@i", MetroTextBox9.Text)
                cmd.Parameters.AddWithValue("@j", ComboBox1.Text)
                cmd.Parameters.AddWithValue("@k", MetroTextBox8.Text)
                cmd.Parameters.AddWithValue("@k", suku)
                cmd.ExecuteNonQuery()
                acsconn.Close()

                MsgBox("4")
            Catch ex As Exception
                MsgBox(ErrorToString)
            End Try
            Rowx = datall.NewRow()
            Rowx.Item("BARKOD") = MetroTextBox3.Text
            Rowx.Item("NAMA ITEM") = MetroTextBox4.Text
            Rowx.Item("HARGA") = MetroTextBox5.Text
            Rowx.Item("KUANTITI") = MetroTextBox6.Text
            Rowx.Item("JUMLAH") = MetroTextBox12.Text
            Rowx.Item("TARIKH") = MetroTextBox7.Text
            Rowx.Item("NAMA PEMBEKAL") = MetroTextBox13.Text
            Rowx.Item("ALAMAT PEMBEKAL") = MetroTextBox10.Text
            Rowx.Item("POSKOD") = MetroTextBox9.Text
            Rowx.Item("NEGERI") = ComboBox1.Text
            Rowx.Item("NO.PESANAN") = MetroTextBox8.Text
            Rowx.Item("KUANTITI PESANAN") = MetroTextBox1.Text
            Rowx.Item("KUANTITI TERIMA") = MetroTextBox11.Text
            Rowx.Item("CATATAN") = MetroTextBox2.Text
            Rowx.Item("KUANTITI1") = kuantiti1
            Rowx.Item("JUMLAH1") = jumlah1
            datall.Rows.Add(Rowx)
            ShowDataInLvw1(datall, ListView1)

        End If
        MetroTextBox3.Text = ""
        MetroTextBox4.Text = ""
        MetroTextBox5.Text = ""
        MetroTextBox6.Text = ""
        MetroTextBox12.Text = ""
        borang()
     
    End Sub

    Private Sub MetroButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton6.Click
        If MetroTextBox11.Text = "" Then
            MsgBox("Sila isi ruangan Kuantiti Diterima")
        Else
            a = Convert.ToDecimal(MetroTextBox5.Text)
            b = Convert.ToDecimal(MetroTextBox6.Text)
            c = Convert.ToDecimal(MetroTextBox11.Text)
            kuantiti = c
            kuantiti1 = b + c
            jumlah = c * a
            jumlah1 = a * (b + c)
            MetroTextBox12.Text = jumlah.ToString("N2")
        End If
    End Sub

    Private Sub MetroButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton4.Click
        For row As Integer = 0 To datall.Rows.Count - 1
            Dim acsconn As System.Data.OleDb.OleDbConnection
            acsconn = New System.Data.OleDb.OleDbConnection
            acsconn.ConnectionString = My.Settings.manageConnectionString
            Dim cmd As New OleDb.OleDbCommand
            Try
                acsconn.Open()
                cmd.Connection = acsconn
                cmd.CommandText = "insert into tambah(kod,nama,harga,kuantiti,tarikh,jumlah,namapembekal,alamatpembekal,poskod,negeri,pesanan,kuantitipesan,kuantititerima,suku) values(@a,@b,@c,@d,@e,@f,@g,@h,@i,@j,@k,@l,@m,@n);"
                cmd.Parameters.AddWithValue("@a", datall.Rows(row).Item("BARKOD"))
                cmd.Parameters.AddWithValue("@b", datall.Rows(row).Item("NAMA ITEM"))
                cmd.Parameters.AddWithValue("@c", datall.Rows(row).Item("HARGA"))
                cmd.Parameters.AddWithValue("@d", datall.Rows(row).Item("KUANTITI PESANAN"))
                cmd.Parameters.AddWithValue("@e", datall.Rows(row).Item("TARIKH"))
                cmd.Parameters.AddWithValue("@f", datall.Rows(row).Item("JUMLAH"))
                cmd.Parameters.AddWithValue("@g", datall.Rows(row).Item("NAMA PEMBEKAL"))
                cmd.Parameters.AddWithValue("@h", datall.Rows(row).Item("ALAMAT PEMBEKAL"))
                cmd.Parameters.AddWithValue("@i", datall.Rows(row).Item("POSKOD"))
                cmd.Parameters.AddWithValue("@j", datall.Rows(row).Item("NEGERI"))
                cmd.Parameters.AddWithValue("@k", datall.Rows(row).Item("NO.PESANAN"))
                cmd.Parameters.AddWithValue("@l", datall.Rows(row).Item("KUANTITI PESANAN"))
                cmd.Parameters.AddWithValue("@m", datall.Rows(row).Item("KUANTITI TERIMA"))
                cmd.Parameters.AddWithValue("@n", suku)
                cmd.ExecuteNonQuery()
                acsconn.Close()

                MetroTextBox3.Text = ""
                MetroTextBox4.Text = ""
                MetroTextBox5.Text = ""
                MetroTextBox6.Text = ""
                MetroTextBox7.Text = ""
                MetroTextBox12.Text = ""
                MetroTextBox13.Text = ""
                MetroTextBox10.Text = ""
                MetroTextBox9.Text = ""
                MetroTextBox8.Text = ""
                MetroTextBox1.Text = ""
                MetroTextBox11.Text = ""
                MetroTextBox2.Text = ""
                ComboBox1.Text = ""
                ' MetroTextBox3.Focus()
                ' Me.Close()
            Catch ex As Exception
                MsgBox(ErrorToString)
            End Try
        Next
        printin()
        aioperation()
        Form1.readdatashow(Form1.ListView1)
        Form1.readdatashow(Form1.ListView2)
        Form1.readdatashow(Form1.ListView3)
        Me.Close()
    End Sub

    Private Sub MetroButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton3.Click
        MetroTextBox13.Text = ""
        MetroTextBox10.Text = ""
        MetroTextBox9.Text = ""
        ComboBox1.SelectedIndex = -1
        MetroTextBox8.Text = ""
        MetroTextBox1.Text = ""
        MetroTextBox11.Text = ""
        MetroTextBox2.Text = ""
    End Sub

    Private Sub MetroButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton5.Click
 
        MetroTextBox3.Text = ""
        MetroTextBox4.Text = ""
        MetroTextBox5.Text = ""
        MetroTextBox6.Text = ""
    End Sub


    Public Sub printin()
        PrintPreviewDialog1.Document = PrintDocument1 'PrintPreviewDialog associate with PrintDocument.

        PrintDocument1.DefaultPageSettings.Landscape = True
        DirectCast(DirectCast(PrintPreviewDialog1.Controls(1), ToolStrip).Items(0), ToolStripButton).Enabled = False
        PrintPreviewDialog1.ShowDialog()

        PrintDialog1.Document = PrintDocument1 'PrintDialog associate with PrintDocument.
        'PrintDocument2.DefaultPageSettings.Landscape = True
        If PrintDialog1.ShowDialog() = DialogResult.OK Then

            PrintDocument1.Print()

        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim drawFont As New Font("Arial", 12.5)
        Dim drawFont1 As New Font("Arial", 10)
        Dim drawBrush As New SolidBrush(Color.DarkBlue)
        Dim drawbrush1 As New SolidBrush(Color.Black)
        Dim drawFormat As New StringFormat
        Dim blackPen As New Pen(Color.Black, 3)

        e.Graphics.DrawString("KEW.PS-1", drawFont, drawBrush, 1000.0F, 10.0F, drawFormat)
        e.Graphics.DrawString("BORANG PENGELUARAN BARANG (BTB)", drawFont, drawBrush, 400.0F, 25.0F, drawFormat)
        e.Graphics.DrawString("(Tatacara Pengurusan Stor 43)", drawFont1, drawBrush, 460.0F, 45.0F, drawFormat)

        'kotak 1st
        'penerangan
        e.Graphics.DrawString("Nama dan Alamat Pembekal:", drawFont, drawBrush, 45.0F, 110.0F, drawFormat)
        e.Graphics.DrawString("Butir-butir Bungkusan:", drawFont, drawBrush, 400.0F, 110.0F, drawFormat)
        e.Graphics.DrawString("Butir-butir Penghantaran:", drawFont, drawBrush, 600.0F, 110.0F, drawFormat)
        e.Graphics.DrawString("No. Pesanan Kerajaan:", drawFont, drawBrush, 870.0F, 110.0F, drawFormat)
        e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 870.0F, 160.0F, drawFormat)


        'horizon line
        e.Graphics.DrawLine(blackPen, 40.0F, 100.0F, 1120.0F, 100.0F)
        e.Graphics.DrawLine(blackPen, 40.0F, 250.0F, 1120.0F, 250.0F)
        'e.Graphics.DrawLine(blackPen, 120.0F, 190.0F, 750.0F, 190.0F)
        'e.Graphics.DrawLine(blackPen, 120.0F, 220.0F, 750.0F, 220.0F)
        'e.Graphics.DrawLine(blackPen, 120.0F, 250.0F, 750.0F, 250.0F)
        'e.Graphics.DrawLine(blackPen, 120.0F, 280.0F, 750.0F, 280.0F)

        'vertical line
        e.Graphics.DrawLine(blackPen, 40.0F, 100.0F, 40.0F, 250.0F)
        e.Graphics.DrawLine(blackPen, 395.0F, 100.0F, 395.0F, 250.0F)
        e.Graphics.DrawLine(blackPen, 595.0F, 100.0F, 595.0F, 250.0F)
        e.Graphics.DrawLine(blackPen, 865.0F, 100.0F, 865.0F, 250.0F)
        e.Graphics.DrawLine(blackPen, 1120.0F, 100.0F, 1120.0F, 250.0F)

        '------------------------------------------------------------------------------------------------------------------------------
        'kotak 2nd
        'penerangan
        e.Graphics.DrawString("No.Kod", drawFont, drawBrush, 45.0F, 295.0F, drawFormat)
        e.Graphics.DrawString("Perihal Barang-barang", drawFont, drawBrush, 140.0F, 295.0F, drawFormat)
        e.Graphics.DrawString("Unit Pengukuran", drawFont, drawBrush, 350.0F, 295.0F, drawFormat)
        e.Graphics.DrawString("Kuantiti", drawFont1, drawBrush, 560.0F, 285.0F, drawFormat)
        e.Graphics.DrawString("Dipesan", drawFont1, drawBrush, 510.0F, 310.0F, drawFormat)
        e.Graphics.DrawString("Diterima", drawFont1, drawBrush, 600.0F, 310.0F, drawFormat)
        e.Graphics.DrawString("Harga", drawFont1, drawBrush, 755.0F, 285.0F, drawFormat)
        e.Graphics.DrawString("Seunit", drawFont1, drawBrush, 710.0F, 310.0F, drawFormat)
        e.Graphics.DrawString("Jumlah", drawFont1, drawBrush, 800.0F, 310.0F, drawFormat)
        e.Graphics.DrawString("Catatan", drawFont, drawBrush, 980.0F, 295.0F, drawFormat)
        'horizon line
        e.Graphics.DrawLine(blackPen, 40.0F, 280.0F, 1120.0F, 280.0F)
        e.Graphics.DrawLine(blackPen, 490.0F, 305.0F, 890.0F, 305.0F)
        e.Graphics.DrawLine(blackPen, 40.0F, 330.0F, 1120.0F, 330.0F)
        ' e.Graphics.DrawLine(blackPen, 40.0F, 400.0F, 1120.0F, 400.0F)

        'vertical line
        e.Graphics.DrawLine(blackPen, 40.0F, 280.0F, 40.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 120.0F, 280.0F, 120.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 345.0F, 280.0F, 345.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 490.0F, 280.0F, 490.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 585.0F, 305.0F, 585.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 680.0F, 280.0F, 680.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 780.0F, 305.0F, 780.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 890.0F, 280.0F, 890.0F, 330.0F)
        e.Graphics.DrawLine(blackPen, 1120.0F, 280.0F, 1120.0F, 330.0F)

        Dim i As Integer = 0
        Dim height As Integer = 370.0F
        Dim newpage As Boolean = False
        For row As Integer = 0 To datall.Rows.Count - 1
            e.Graphics.DrawLine(blackPen, 40.0F, 370.0F + i, 1120.0F, height)
            'vertical
            e.Graphics.DrawLine(blackPen, 40.0F, 330.0F, 40.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 120.0F, 280.0F, 120.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 345.0F, 280.0F, 345.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 490.0F, 280.0F, 490.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 585.0F, 305.0F, 585.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 680.0F, 280.0F, 680.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 780.0F, 305.0F, 780.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 890.0F, 280.0F, 890.0F, 370.0F + i)
            e.Graphics.DrawLine(blackPen, 1120.0F, 280.0F, 1120.0F, 370.0F + i)

            'detail
            e.Graphics.DrawString(datall.Rows(row).Item("BARKOD"), drawFont1, drawbrush1, 60.0F, 340.0F + i, drawFormat)
            e.Graphics.DrawString(datall.Rows(row).Item("NAMA ITEM"), drawFont1, drawbrush1, 130.0F, 340.0F + i, drawFormat)
            e.Graphics.DrawString("UNIT", drawFont1, drawbrush1, 400.0F, 340.0F + i, drawFormat)
            e.Graphics.DrawString(datall.Rows(row).Item("KUANTITI PESANAN"), drawFont1, drawbrush1, 500.0F, 340.0F + i, drawFormat)
            e.Graphics.DrawString(datall.Rows(row).Item("KUANTITI TERIMA"), drawFont1, drawbrush1, 595.0F, 340.0F + i, drawFormat)
            e.Graphics.DrawString(datall.Rows(row).Item("HARGA"), drawFont1, drawbrush1, 685.0F, 340.0F + i, drawFormat)
            e.Graphics.DrawString(datall.Rows(row).Item("JUMLAH"), drawFont1, drawbrush1, 785.0F, 340.0F + i, drawFormat)
            e.Graphics.DrawString(datall.Rows(row).Item("CATATAN"), drawFont1, drawbrush1, 895.0F, 340.0F + i, drawFormat)
            i = i + 40
            height = height + 40

            'detail
            e.Graphics.DrawString(datall.Rows(row1).Item("NAMA PEMBEKAL"), drawFont1, drawbrush1, 45.0F, 130.0F, drawFormat)
            e.Graphics.DrawString(datall.Rows(row1).Item("ALAMAT PEMBEKAL"), drawFont1, drawbrush1, 45.0F, 150.0F, drawFormat)
            e.Graphics.DrawString(datall.Rows(row1).Item("POSKOD"), drawFont1, drawbrush1, 45.0F, 170.0F, drawFormat)
            e.Graphics.DrawString(datall.Rows(row1).Item("NO.PESANAN"), drawFont1, drawbrush1, 870.0F, 130.0F, drawFormat)
            e.Graphics.DrawString(datall.Rows(row1).Item("TARIKH"), drawFont1, drawbrush1, 870.0F, 190.0F, drawFormat)
            If row > 5 Then
                newpage = True
            End If
        Next

        '-------------------------------------------------------------------------------------------------------------------------------
        If newpage Then

            'kotak 3rd
            'penerangan
            e.Graphics.DrawString("...............................", drawFont, drawBrush, 45.0F, height + 80, drawFormat)
            e.Graphics.DrawString("(Tandatangan Pegawai Penerima)", drawFont, drawBrush, 45.0F, height + 100, drawFormat)
            e.Graphics.DrawString("Nama:", drawFont, drawBrush, 45.0F, height + 130, drawFormat)
            e.Graphics.DrawString("Jawatan:", drawFont, drawBrush, 45.0F, height + 160, drawFormat)
            e.Graphics.DrawString("Jabatan:", drawFont, drawBrush, 45.0F, height + 190, drawFormat)
            e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 45.0F, height + 220, drawFormat)

            e.Graphics.DrawString("...............................", drawFont, drawBrush, 400.0F, height + 80, drawFormat)
            e.Graphics.DrawString("(Tandatangan Pegawai Pengesah)", drawFont, drawBrush, 400.0F, height + 100, drawFormat)
            e.Graphics.DrawString("Nama:", drawFont, drawBrush, 400.0F, height + 130, drawFormat)
            e.Graphics.DrawString("Jawatan:", drawFont, drawBrush, 400.0F, height + 160, drawFormat)
            e.Graphics.DrawString("Jabatan:", drawFont, drawBrush, 400.0F, height + 190, drawFormat)
            e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 400.0F, height + 220, drawFormat)

            e.Graphics.DrawString("Nota:", drawFont, drawBrush, 755.0F, height + 10, drawFormat)

            'horizon line
            e.Graphics.DrawLine(blackPen, 40.0F, height, 1120.0F, height)
            e.Graphics.DrawLine(blackPen, 40.0F, height + 250.0F, 1120.0F, height + 250.0F)

            'vertical line
            e.Graphics.DrawLine(blackPen, 40.0F, height, 40.0F, height + 250.0F)
            e.Graphics.DrawLine(blackPen, 395.0F, height, 395.0F, height + 250.0F)
            e.Graphics.DrawLine(blackPen, 750.0F, height, 750.0F, height + 250.0F)
            e.Graphics.DrawLine(blackPen, 1120.0F, height, 1120.0F, height + 250.0F)
            ' e.HasMorePages = True
            ' e.HasMorePages = False
        Else

            'kotak 3rd
            'penerangan
            e.Graphics.DrawString("...............................", drawFont, drawBrush, 45.0F, height + 80, drawFormat)
            e.Graphics.DrawString("(Tandatangan Pegawai Penerima)", drawFont, drawBrush, 45.0F, height + 100, drawFormat)
            e.Graphics.DrawString("Nama:", drawFont, drawBrush, 45.0F, height + 130, drawFormat)
            e.Graphics.DrawString("Jawatan:", drawFont, drawBrush, 45.0F, height + 160, drawFormat)
            e.Graphics.DrawString("Jabatan:", drawFont, drawBrush, 45.0F, height + 190, drawFormat)
            e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 45.0F, height + 220, drawFormat)

            e.Graphics.DrawString("...............................", drawFont, drawBrush, 400.0F, height + 80, drawFormat)
            e.Graphics.DrawString("(Tandatangan Pegawai Pengesah)", drawFont, drawBrush, 400.0F, height + 100, drawFormat)
            e.Graphics.DrawString("Nama:", drawFont, drawBrush, 400.0F, height + 130, drawFormat)
            e.Graphics.DrawString("Jawatan:", drawFont, drawBrush, 400.0F, height + 160, drawFormat)
            e.Graphics.DrawString("Jabatan:", drawFont, drawBrush, 400.0F, height + 190, drawFormat)
            e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 400.0F, height + 220, drawFormat)

            e.Graphics.DrawString("Nota:", drawFont, drawBrush, 755.0F, height + 10, drawFormat)

            'horizon line
            e.Graphics.DrawLine(blackPen, 40.0F, height, 1120.0F, height)
            e.Graphics.DrawLine(blackPen, 40.0F, height + 250.0F, 1120.0F, height + 250.0F)

            'vertical line
            e.Graphics.DrawLine(blackPen, 40.0F, height, 40.0F, height + 250.0F)
            e.Graphics.DrawLine(blackPen, 395.0F, height, 395.0F, height + 250.0F)
            e.Graphics.DrawLine(blackPen, 750.0F, height, 750.0F, height + 250.0F)
            e.Graphics.DrawLine(blackPen, 1120.0F, height, 1120.0F, height + 250.0F)
            ' e.HasMorePages = False
        End If
      


      


        'e.Graphics.DrawString("asdasdadsasdasd", drawFont1, drawbrush1, 60.0F, 340.0F, drawFormat)
        'e.Graphics.DrawString(MetroTextBox4.Text, drawFont1, drawbrush1, 130.0F, 350.0F, drawFormat)
        'e.Graphics.DrawString("UNIT", drawFont1, drawbrush1, 400.0F, 350.0F, drawFormat)
        'e.Graphics.DrawString(datall.Rows(row1).Item("BILPESAN"), drawFont1, drawbrush1, 500.0F, 350.0F, drawFormat)
        'e.Graphics.DrawString(datall.Rows(row1).Item("CATATAN"), drawFont1, drawbrush1, 895.0F, 350.0F, drawFormat)

        'If dahade Then
        '    e.Graphics.DrawString(g.ToString("N2"), drawFont1, drawbrush1, 685.0F, 350.0F, drawFormat)
        '    e.Graphics.DrawString(b.ToString, drawFont1, drawbrush1, 595.0F, 350.0F, drawFormat)

        'Else
        '    e.Graphics.DrawString(d.ToString("N2"), drawFont1, drawbrush1, 685.0F, 350.0F, drawFormat)
        '    e.Graphics.DrawString(f.ToString, drawFont1, drawbrush1, 595.0F, 350.0F, drawFormat)

        'End If
        ' e.Graphics.DrawString(totalsum.ToString("N2"), drawFont1, drawbrush1, 785.0F, 350.0F, drawFormat)
    End Sub

    'Private Sub MetroButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton7.Click
    '    printin()
    '    'Dim thisDate As Date
    '    'Dim thisMonth As Integer
    '    '' thisDate = Date.Now.ToString
    '    'thisDate = MetroTextBox7.Text
    '    'thisMonth = Month(thisDate)
    '    'MsgBox(thisMonth.ToString)
    '    'For row As Integer = 0 To datall.Rows.Count - 1
    '    '    'ds.Tables.Clear()
    '    '    Dim oleConn As System.Data.OleDb.OleDbConnection
    '    '    Dim da As OleDb.OleDbDataAdapter
    '    '    ds = New DataSet
    '    '    oleConn = New System.Data.OleDb.OleDbConnection
    '    '    oleConn.ConnectionString = My.Settings.manageConnectionString
    '    '    Try

    '    '        oleConn.Open()
    '    '        da = New OleDb.OleDbDataAdapter("select nama from detail where kod like '%" & datall.Rows(row).Item("BARKOD") & "%'", oleConn)
    '    '        da.Fill(ds, "infox")
    '    '        oleConn.Close()

    '    '        'MetroTextBox11.Text = ds.Tables("infox").Rows(row)("nama")
    '    '        'MetroTextBox10.Text = ds.Tables("infox").Rows(row)("harga")
    '    '        'MetroTextBox9.Text = ds.Tables("infox").Rows(row)("kuantiti")
    '    '        'MetroTextBox8.Text = ds.Tables("infox").Rows(row)("tarikh")
    '    '        'MetroLabel16.Visible = True
    '    '        'MetroTextBox15.Visible = True
    '    '        dahade = True
    '    '        MsgBox(ds.Tables("infox").Rows(row)("nama").ToString)
    '    '        MsgBox("2")
    '    '    Catch ex As Exception
    '    '        MsgBox("Item Tiada Dalam Simpanan. Sila Masukkan Item Mengikut Seksyen Item Masuk Dahulu.")
    '    '        dahade = False
    '    '    End Try
    '    '    ds.Tables.Clear()
    '    ' aioperation(row)
    '    'Next
    'End Sub
  
    Public Sub aioperation()
        For row As Integer = 0 To datall.Rows.Count - 1

            Dim acsconn As System.Data.OleDb.OleDbConnection
            acsconn = New System.Data.OleDb.OleDbConnection
            acsconn.ConnectionString = My.Settings.manageConnectionString
            Dim cmd As New OleDb.OleDbCommand
            Try
                acsconn.Open()
                cmd.Connection = acsconn
                cmd.CommandText = "update detail set kod=@a,nama=@b,harga=@c,kuantiti=@d,tarikh=@e,jumlah=@f,namapembekal=@g,alamatpembekal=@h,poskod=@i,pesanan=@k,negeri=@j,suku=@l where kod = @a;"
                cmd.Parameters.AddWithValue("@a", datall.Rows(row).Item("BARKOD"))
                cmd.Parameters.AddWithValue("@b", datall.Rows(row).Item("NAMA ITEM"))
                cmd.Parameters.AddWithValue("@c", datall.Rows(row).Item("HARGA"))
                cmd.Parameters.AddWithValue("@d", datall.Rows(row).Item("KUANTITI1"))
                cmd.Parameters.AddWithValue("@e", datall.Rows(row).Item("TARIKH"))
                cmd.Parameters.AddWithValue("@f", datall.Rows(row).Item("JUMLAH1"))
                cmd.Parameters.AddWithValue("@g", datall.Rows(row).Item("NAMA PEMBEKAL"))
                cmd.Parameters.AddWithValue("@h", datall.Rows(row).Item("ALAMAT PEMBEKAL"))
                cmd.Parameters.AddWithValue("@i", datall.Rows(row).Item("POSKOD"))
                cmd.Parameters.AddWithValue("@j", datall.Rows(row).Item("NEGERI"))
                cmd.Parameters.AddWithValue("@k", datall.Rows(row).Item("NO.PESANAN"))
                cmd.Parameters.AddWithValue("@l", suku)
                cmd.ExecuteNonQuery()
                acsconn.Close()
                MsgBox("berjaya")
            Catch ex As Exception
                MsgBox("gagal ubah")
            End Try
        Next
        
    End Sub

    Private Sub MetroButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton8.Click
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Try
            oleConn.Open()
            da = New OleDb.OleDbDataAdapter("select * from detail where kod like '%" & MetroTextBox3.Text & "%'", oleConn)
            da.Fill(ds, "infox")
            oleConn.Close()

            MetroTextBox4.Text = ds.Tables("infox").Rows(0)("nama")
            MetroTextBox5.Text = ds.Tables("infox").Rows(0)("harga")
            MetroTextBox6.Text = ds.Tables("infox").Rows(0)("kuantiti")
            'a = ds.Tables("infox").Rows(row)("kuantiti")
            'harga = ds.Tables("infox").Rows(row)("harga")
            MetroTextBox13.Text = ds.Tables("infox").Rows(0)("namapembekal")
            MetroTextBox10.Text = ds.Tables("infox").Rows(0)("alamatpembekal")
            MetroTextBox9.Text = ds.Tables("infox").Rows(0)("poskod")
            ComboBox1.Text = ds.Tables("infox").Rows(0)("negeri")
            MetroTextBox8.Text = ds.Tables("infox").Rows(0)("pesanan")
            'MetroTextBox1.Text = ds.Tables("infox").Rows(0)("kuantitipesan")
            'MetroTextBox11.Text = ds.Tables("infox").Rows(0)("kuantititerima")
            ' MetroTextBox8.Text = ds.Tables("infox").Rows(row)("tarikh")
            dahade = True
        Catch ex As Exception
            MsgBox("Item Tiada Dalam Simpanan. Sila Masukkan Item Mengikut Seksyen Item Masuk Dahulu.")
            dahade = False
            'MetroTextBox1.Text = ""
            'MetroTextBox2.Text = ""
            'MetroTextBox3.Text = ""
            ' MetroTextBox8.Text = ""
            ' MetroLabel16.Visible = False
            'MetroTextBox15.Visible = False
            ' MetroTextBox12.Text = ""
            ' MetroTextBox12.Focus()
            'dahade = False
        End Try
    End Sub
    Public Sub borang()
        Dim jumlah As Decimal = 0
        For row As Integer = 0 To datall.Rows.Count - 1
            jumlah += datall.Rows(row).Item("HARGA")
        Next
        MetroTextBox14.Text = "RM" + jumlah.ToString("N2")
    End Sub

    Private Sub MetroButton9_Click(sender As Object, e As EventArgs)
        borang()
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        datall.Clear()
        Me.Close()
    End Sub

    Private Sub MetroButton7_Click(sender As Object, e As EventArgs)
        printin()
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        row1 = ListView1.SelectedItems(0).Index

        datall.Rows(row1).Delete()
        ShowDataInLvw1(datall, ListView1)
        'MsgBox(datall.Rows(row1).Item("BARKOD"))
    End Sub
End Class