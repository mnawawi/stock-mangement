Imports System.Data.OleDb

Public Class Form1
    Public Shared ds As DataSet
    Dim row As Integer
    Dim row1 As Integer
    Dim jumlah1 As Integer
    Dim totalsum As Decimal = 0
    Dim format1 As String = "dd/MM/yyyy"
    Dim today As String
    Dim dahade As Boolean = False
    Dim a As Decimal = 0
    Dim b As Decimal = 0
    Dim a1 As Decimal = 0
    Dim b1 As Decimal = 0
    Dim sum As Decimal = 0
    Dim c As Decimal = 0
    Dim d As Decimal = 0
    Dim f As Decimal = 0
    Dim g As Decimal = 0
    Dim h As Decimal = 0
    Dim j As Decimal = 0
    Dim k As Decimal = 0
    Dim minus As Decimal = 0
    Public Shared datall As DataTable
    Public Shared detail As Boolean = False
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        dahade = False
        MetroTextBox1.Text = ""
        MetroTextBox2.Text = ""
        MetroButton2.Visible = False
        MetroTabControl1.Visible = False
        readdatashow(ListView1)
        readdatashow(ListView2)
        readdatashow(ListView3)
        updateprice()
        today = DateTime.Now.ToString(format1)
        Label2.Text = today

        Timer1.Start()
        table1()
        stok(ListView4)
        kedudukan(ListView5)
        'MetroLabel16.Visible = False
        'MetroTextBox15.Visible = False
    End Sub
    Public Sub table1()
        datall = New DataTable("DETAIL")

        Dim column1 As DataColumn = New DataColumn("NAMAPEMBEKAL")
        column1.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column1)

        Dim column2 As DataColumn = New DataColumn("ALAMATPEMBEKAL")
        column2.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column2)

        Dim column3 As DataColumn = New DataColumn("POSKOD")
        column3.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column3)

        Dim column4 As DataColumn = New DataColumn("PESANAN")
        column4.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column4)

        Dim column5 As DataColumn = New DataColumn("CATATAN")
        column5.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column5)

        Dim column6 As DataColumn = New DataColumn("BILPESAN")
        column6.DataType = System.Type.GetType("System.String")
        datall.Columns.Add(column6)

        Dim row1 As DataRow
        row1 = datall.NewRow()
        row1.Item("NAMAPEMBEKAL") = ""
        row1.Item("ALAMATPEMBEKAL") = ""
        row1.Item("POSKOD") = ""
        row1.Item("PESANAN") = ""
        row1.Item("CATATAN") = ""
        row1.Item("BILPESAN") = ""
        datall.Rows.Add(row1)
    End Sub
    Private Sub TabControlAction(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If MetroTabControl1.SelectedTab.Text = "ITEM MASUK" Then

                readdatashow(ListView1)

            ElseIf MetroTabControl1.SelectedTab.Text = "ITEM KELUAR" Then
                readdatashow(ListView2)
                'MetroTextBox12.Focus()
            ElseIf MetroTabControl1.SelectedTab.Text = "SENARAI ITEM" Then
                readdatashow(ListView3)

            ElseIf MetroTabControl1.SelectedTab.Text = "KAWALAN STOK" Then
                stok(ListView4)
                kedudukan(ListView5)
            End If
        Catch ex As Exception
            MsgBox(ErrorToString)
        End Try


    End Sub

    Private Sub MetroButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton1.Click

        If MetroTextBox1.Text = "" Or MetroTextBox2.Text = "" Then
            MsgBox("Sila Masukkan ID Pengguna Atau Kata Laluan")
        ElseIf MetroTextBox1.Text = "admin" And MetroTextBox2.Text = "admin" Then
            MetroTabControl1.Visible = True
            MetroButton2.Visible = True
            MetroTextBox1.Text = ""
            MetroTextBox2.Text = ""
           
            MetroLabel4.Text = "MASTER USER"
            MetroButton10.Text = "LOG KELUAR"
        Else
            Dim oleConn As System.Data.OleDb.OleDbConnection
            oleConn = New System.Data.OleDb.OleDbConnection
            Dim da As New OleDb.OleDbDataAdapter
            ds = New DataSet
            oleConn.ConnectionString = My.Settings.manageConnectionString
            Dim command As New OleDbCommand("SELECT [namastaf],nokp FROM [staff] WHERE [username] = username AND [katalaluan] = password", oleConn)
            Dim usernameParam As New OleDbParameter("username", MetroTextBox1.Text)
            Dim passwordParam As New OleDbParameter("password", MetroTextBox2.Text)
            command.Parameters.Add(usernameParam)
            command.Parameters.Add(passwordParam)
            command.Connection.Open()
            da.SelectCommand = command
            da.Fill(ds, "login")
            'MsgBox(ds.Tables("login").Rows(row)("namastaf"))
            Dim reader As OleDbDataReader = command.ExecuteReader()
            If reader.HasRows Then
                'If MetroTextBox1.Text = ds.Tables("login").Rows(row)("username") And MetroTextBox2.Text = ds.Tables("login").Rows(row)("katalaluan") Then
                MetroTabControl1.Visible = True
                MetroLabel4.Text = ds.Tables("login").Rows(row)("namastaf")
                'MsgBox(ds.Tables("login").Rows(row).Item("nokp"))
                'End If

            Else
                MetroTabControl1.Visible = False
                MessageBox.Show("salah id atau katalaluan", "makluman")
            End If
            command.Connection.Close()
            updatelogin()

            MetroTextBox1.Text = ""
            MetroTextBox2.Text = ""
            'MetroTextBox1.Enabled = False
            'MetroTextBox2.Enabled = False
            MetroButton10.Text = "LOG KELUAR"
        End If

        If MetroButton10.Text = "LOG KELUAR" Then
            MetroTextBox1.Enabled = False
            MetroTextBox2.Enabled = False
            MetroButton1.Enabled = False
       
        End If
    End Sub

    Public Sub updatelogin()
        Dim acsconn As System.Data.OleDb.OleDbConnection
        acsconn = New System.Data.OleDb.OleDbConnection
        acsconn.ConnectionString = My.Settings.manageConnectionString
        Dim cmd As New OleDb.OleDbCommand
        Try
            acsconn.Open()
            cmd.Connection = acsconn
            cmd.CommandText = "update staff set tarikh=@a,masa=@b where nokp = @c;"
            cmd.Parameters.AddWithValue("@a", Label2.Text)
            cmd.Parameters.AddWithValue("@b", Label1.Text + " " + Label3.Text)
            cmd.Parameters.AddWithValue("@c", ds.Tables("login").Rows(row).Item("nokp"))

            cmd.ExecuteNonQuery()
            acsconn.Close()
            MsgBox("berjaya")
        Catch ex As Exception
            MsgBox("gagal ubah")
        End Try
    End Sub

    Public Sub readdatashow(ByVal lvw As ListView)

        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        oleConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT * FROM detail", oleConn)
        da.Fill(ds, "detail")
        oleConn.Close()
        lvw.Clear()
        With lvw
            .Items.Clear()
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .Columns.Add("ID", 0, HorizontalAlignment.Left)
            .Columns.Add("KOD", 100, HorizontalAlignment.Left)
            .Columns.Add("NAMA ITEM", 230, HorizontalAlignment.Left)
            .Columns.Add("HARGA (RM)", 100, HorizontalAlignment.Left)
            .Columns.Add("KUANTITI", 80, HorizontalAlignment.Left)
            .Columns.Add("JUMLAH (RM)", 100, HorizontalAlignment.Left)
            .Columns.Add("TARIKH", 130, HorizontalAlignment.Left)


        End With
        For Each row As DataRow In ds.Tables("detail").Rows
            Dim lst As ListViewItem
            lst = lvw.Items.Add(row(0))
            For i As Integer = 1 To ds.Tables("detail").Columns.Count - 1
                lst.SubItems.Add(row(i))
            Next
        Next

        'row1 = 0
        'jumlah1 = 0
        'For Each huhu As DataRow In ds.Tables("infox").Rows
        '    jumlah1 = jumlah1 + ds.Tables("infox").Rows(row1)("jumlah")
        '    row1 = row + 1
        '    MsgBox(jumlah1)
        'Next

    End Sub
    Private Sub updateprice()
        jumlah1 = 0
        For i As Integer = 0 To ds.Tables("detail").Rows.Count - 1
            jumlah1 += ds.Tables("detail").Rows(i).Item("jumlah")
        Next
        MetroTextBox13.Text = jumlah1.ToString("N2")
        MetroTextBox16.Text = jumlah1.ToString("N2")
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'The below is desined for details view.
        If ListView1.View = View.Details Then
            PrintDetails(e)
        End If
    End Sub

    Private Sub PrintDetails(ByRef e As System.Drawing.Printing.PrintPageEventArgs)
        Dim drawFont As New Font("Arial", 26)
        Dim drawBrush As New SolidBrush(Color.DarkBlue)
        Dim drawbrush1 As New SolidBrush(Color.Black)
        Dim drawFormat As New StringFormat
        Dim blackPen As New Pen(Color.Black, 3)
        e.Graphics.DrawString("SENARAI ITEM", drawFont, drawBrush, 280.0F, 40.0F, drawFormat)

        Static LastIndex As Integer = 0
        Static CurrentPage As Integer = 0
        'Getting the current dpi so the textleftpad 
        'will be the same on a different dpi than 
        'the 96 i'm using.  Won't make much of a difference though.
        Dim DpiGraphics As Graphics = Me.CreateGraphics
        Dim DpiX As Integer = DpiGraphics.DpiX
        Dim DpiY As Integer = DpiGraphics.DpiY
        DpiGraphics.Dispose()
        Dim X, Y As Integer
        Dim ImageWidth As Integer
        Dim TextRect As Rectangle = Rectangle.Empty
        Dim TextLeftPad As Single = CSng(4 * (DpiX / 96)) '4 pixel pad on the left.
        Dim ColumnHeaderHeight As Single = CSng(ListView1.Font.Height + (10 * (DpiX / 96))) '5 pixel pad on the top an bottom
        Dim StringFormat As New StringFormat
        Dim PageNumberWidth As Single = e.Graphics.MeasureString(CStr(CurrentPage), ListView1.Font).Width
        'Specify the text should be drawn in the center of the line and
        'that the text should not be wrapped and the text should show
        'ellipsis would cut off.
        StringFormat.FormatFlags = StringFormatFlags.NoWrap
        StringFormat.Trimming = StringTrimming.EllipsisCharacter
        StringFormat.LineAlignment = StringAlignment.Center
        CurrentPage += 1
        'Start the x and  y at the top left margin.
        X = CInt(e.MarginBounds.X)
        Y = CInt(e.MarginBounds.Y)
        'Draw the column headers
        For ColumnIndex As Integer = 0 To ListView1.Columns.Count - 1
            TextRect.X = X
            TextRect.Y = Y
            TextRect.Width = ListView1.Columns(ColumnIndex).Width
            TextRect.Height = ColumnHeaderHeight
            e.Graphics.FillRectangle(Brushes.LightGray, TextRect)
            e.Graphics.DrawRectangle(Pens.DarkGray, TextRect)
            'TextLeftPad adds a little padding from the gridline.
            'Add it to the left and subtract it from the right.
            TextRect.X += TextLeftPad
            TextRect.Width -= TextLeftPad
            e.Graphics.DrawString(ListView1.Columns(ColumnIndex).Text, ListView1.Font, Brushes.Black, TextRect, StringFormat)
            'Move the x position over the width of the column width.
            'Since I subtracted some padding add the padding back
            'when offsetting.
            X += TextRect.Width + TextLeftPad
        Next
        'Just drew the headers.  Move the Y down the height
        'of the column headers.
        Y += ColumnHeaderHeight
        'Now draw the items.  If this is the first page then the 
        'last index will be zero.  If its not then the last index
        'will be the last index we tried to draw but had no room.
        For i = LastIndex To ListView1.Items.Count - 1
            With ListView1.Items(i)
                'Start the x at the pages left margin.
                X = CInt(e.MarginBounds.X)
                'Check for Last Line
                If Y + .Bounds.Height > e.MarginBounds.Bottom Then
                    'This item won't fit.
                    'subtract 1 from i so the next time this sub
                    'is entered we can start with this item.
                    LastIndex = i - 1
                    e.HasMorePages = True
                    StringFormat.Dispose()
                    'Draw the current page number before leaving.
                    e.Graphics.DrawString(CStr(CurrentPage), ListView1.Font, Brushes.Black, (e.PageBounds.Width - PageNumberWidth) / 2, e.PageBounds.Bottom - ListView1.Font.Height * 2)
                    Exit Sub
                End If
                'Print Images.
                'The image width is used so we can draw the gridline
                'around the image about to be drawn.  You'll see it 
                'below.
                ImageWidth = 0
                If ListView1.SmallImageList IsNot Nothing Then
                    'If the image key is set then draw the image
                    'with the key .  If not draw the image with the
                    'index.  A tiny bit of validation would be good.
                    If Not String.IsNullOrEmpty(.ImageKey) Then
                        e.Graphics.DrawImage(ListView1.SmallImageList.Images(.ImageKey), X, Y)
                    ElseIf .ImageIndex >= 0 Then
                        e.Graphics.DrawImage(ListView1.SmallImageList.Images(.ImageIndex), X, Y)
                    End If
                    ImageWidth = ListView1.SmallImageList.ImageSize.Width
                End If
                'Now draw the subitems.  using the columns count so the 
                'grid lines can be drawn.  If used the subitems count then
                'the table would not be full if some subitems where less
                'than others.
                For ColumnIndex As Integer = 0 To ListView1.Columns.Count - 1
                    TextRect.X = X
                    TextRect.Y = Y
                    TextRect.Width = ListView1.Columns(ColumnIndex).Width
                    TextRect.Height = .Bounds.Height
                    If ListView1.GridLines Then
                        e.Graphics.DrawRectangle(Pens.DarkGray, TextRect)
                    End If
                    'If an image is drawn then shift over the x to 
                    'accomadate its width. If this was shifted before
                    'now then the gridline with draw rect above would be
                    ' on the wrong side of the image.
                    If ColumnIndex = 0 Then TextRect.X += ImageWidth
                    'Add a little padding from the gridline.
                    TextRect.X += TextLeftPad
                    TextRect.Width -= TextLeftPad
                    If ColumnIndex < .SubItems.Count Then
                        'This item has at least the same number of
                        'subitems as the current column index.
                        e.Graphics.DrawString(.SubItems(ColumnIndex).Text, ListView1.Font, Brushes.Black, TextRect, StringFormat)
                    End If
                    'Shift the x of the width of this subitem.
                    'Add some padding to the left side of the text
                    'so need to add it back.
                    X += TextRect.Width + TextLeftPad
                Next
                'Set the next line
                Y += .Bounds.Height
            End With
        Next
        'Draw the final page number.
        e.Graphics.DrawString(CStr(CurrentPage), ListView1.Font, Brushes.Black, (e.PageBounds.Width - PageNumberWidth) / 2, e.PageBounds.Bottom - ListView1.Font.Height * 2)
        StringFormat.Dispose()
        LastIndex = 0
        CurrentPage = 0
    End Sub



    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label1.Text = Format(Now, "hh:mm:ss")
        If Date.Now.Hour < 12 Then
            Label3.Text = "AM"
        Else
            Label3.Text = "PM"
        End If
    End Sub


    Private Sub MetroButton10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton10.Click
        If MetroButton10.Text = "LOG KELUAR" Then
            MetroTabControl1.Visible = False
            MetroTextBox1.Enabled = True
            MetroTextBox2.Enabled = True
            MetroButton2.Visible = False
            MetroButton1.Enabled = True
            MetroButton10.Text = "KELUAR"
            MetroLabel4.Text = ""
        ElseIf MetroButton10.Text = "KELUAR" Then
            Me.Close()
        End If
    End Sub
    Private Sub printin()
        PrintPreviewDialog1.Document = PrintDocument2 'PrintPreviewDialog associate with PrintDocument.

        PrintDocument2.DefaultPageSettings.Landscape = True
        DirectCast(DirectCast(PrintPreviewDialog1.Controls(1), ToolStrip).Items(0), ToolStripButton).Enabled = False
        PrintPreviewDialog1.ShowDialog()

        PrintDialog2.Document = PrintDocument2 'PrintDialog associate with PrintDocument.
        'PrintDocument2.DefaultPageSettings.Landscape = True
        If PrintDialog2.ShowDialog() = DialogResult.OK Then

            PrintDocument2.Print()

        End If
    End Sub

    'Private Sub PrintDocument2_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage
    '    Dim drawFont As New Font("Arial", 12.5)
    '    Dim drawFont1 As New Font("Arial", 10)
    '    Dim drawBrush As New SolidBrush(Color.DarkBlue)
    '    Dim drawbrush1 As New SolidBrush(Color.Black)
    '    Dim drawFormat As New StringFormat
    '    Dim blackPen As New Pen(Color.Black, 3)

    '    e.Graphics.DrawString("BORANG PENGELUARAN BARANG (BTB)", drawFont, drawBrush, 400.0F, 25.0F, drawFormat)
    '    e.Graphics.DrawString("(Tatacara Pengurusan Stor 43)", drawFont1, drawBrush, 460.0F, 45.0F, drawFormat)

    '    'kotak 1st
    '    'penerangan
    '    e.Graphics.DrawString("Nama dan Alamat Pembekal:", drawFont, drawBrush, 45.0F, 110.0F, drawFormat)
    '    e.Graphics.DrawString("Butir-butir Bungkusan:", drawFont, drawBrush, 400.0F, 110.0F, drawFormat)
    '    e.Graphics.DrawString("Butir-butir Penghantaran:", drawFont, drawBrush, 600.0F, 110.0F, drawFormat)
    '    e.Graphics.DrawString("No. Pesanan Kerajaan:", drawFont, drawBrush, 870.0F, 110.0F, drawFormat)
    '    e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 870.0F, 160.0F, drawFormat)


    '    'horizon line
    '    e.Graphics.DrawLine(blackPen, 40.0F, 100.0F, 1120.0F, 100.0F)
    '    e.Graphics.DrawLine(blackPen, 40.0F, 250.0F, 1120.0F, 250.0F)
    '    'e.Graphics.DrawLine(blackPen, 120.0F, 190.0F, 750.0F, 190.0F)
    '    'e.Graphics.DrawLine(blackPen, 120.0F, 220.0F, 750.0F, 220.0F)
    '    'e.Graphics.DrawLine(blackPen, 120.0F, 250.0F, 750.0F, 250.0F)
    '    'e.Graphics.DrawLine(blackPen, 120.0F, 280.0F, 750.0F, 280.0F)

    '    'vertical line
    '    e.Graphics.DrawLine(blackPen, 40.0F, 100.0F, 40.0F, 250.0F)
    '    e.Graphics.DrawLine(blackPen, 395.0F, 100.0F, 395.0F, 250.0F)
    '    e.Graphics.DrawLine(blackPen, 595.0F, 100.0F, 595.0F, 250.0F)
    '    e.Graphics.DrawLine(blackPen, 865.0F, 100.0F, 865.0F, 250.0F)
    '    e.Graphics.DrawLine(blackPen, 1120.0F, 100.0F, 1120.0F, 250.0F)

    '    '------------------------------------------------------------------------------------------------------------------------------
    '    'kotak 2nd
    '    'penerangan
    '    e.Graphics.DrawString("No.Kod", drawFont, drawBrush, 45.0F, 295.0F, drawFormat)
    '    e.Graphics.DrawString("Perihal Barang-barang", drawFont, drawBrush, 140.0F, 295.0F, drawFormat)
    '    e.Graphics.DrawString("Unit Pengukuran", drawFont, drawBrush, 350.0F, 295.0F, drawFormat)
    '    e.Graphics.DrawString("Kuantiti", drawFont1, drawBrush, 560.0F, 285.0F, drawFormat)
    '    e.Graphics.DrawString("Dipesan", drawFont1, drawBrush, 510.0F, 310.0F, drawFormat)
    '    e.Graphics.DrawString("Diterima", drawFont1, drawBrush, 600.0F, 310.0F, drawFormat)
    '    e.Graphics.DrawString("Harga", drawFont1, drawBrush, 755.0F, 285.0F, drawFormat)
    '    e.Graphics.DrawString("Seunit", drawFont1, drawBrush, 710.0F, 310.0F, drawFormat)
    '    e.Graphics.DrawString("Jumlah", drawFont1, drawBrush, 800.0F, 310.0F, drawFormat)
    '    e.Graphics.DrawString("Catatan", drawFont, drawBrush, 980.0F, 295.0F, drawFormat)
    '    'horizon line
    '    e.Graphics.DrawLine(blackPen, 40.0F, 280.0F, 1120.0F, 280.0F)
    '    e.Graphics.DrawLine(blackPen, 490.0F, 305.0F, 890.0F, 305.0F)
    '    e.Graphics.DrawLine(blackPen, 40.0F, 330.0F, 1120.0F, 330.0F)
    '    e.Graphics.DrawLine(blackPen, 40.0F, 400.0F, 1120.0F, 400.0F)

    '    'vertical line
    '    e.Graphics.DrawLine(blackPen, 40.0F, 280.0F, 40.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 120.0F, 280.0F, 120.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 345.0F, 280.0F, 345.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 490.0F, 280.0F, 490.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 585.0F, 305.0F, 585.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 680.0F, 280.0F, 680.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 780.0F, 305.0F, 780.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 890.0F, 280.0F, 890.0F, 400.0F)
    '    e.Graphics.DrawLine(blackPen, 1120.0F, 280.0F, 1120.0F, 400.0F)
    '    '-------------------------------------------------------------------------------------------------------------------------------
    '    'kotak 3rd
    '    'penerangan
    '    e.Graphics.DrawString("...............................", drawFont, drawBrush, 45.0F, 510.0F, drawFormat)
    '    e.Graphics.DrawString("(Tandatangan Pegawai Penerima)", drawFont, drawBrush, 45.0F, 530.0F, drawFormat)
    '    e.Graphics.DrawString("Nama:", drawFont, drawBrush, 45.0F, 560.0F, drawFormat)
    '    e.Graphics.DrawString("Jawatan:", drawFont, drawBrush, 45.0F, 590.0F, drawFormat)
    '    e.Graphics.DrawString("Jabatan:", drawFont, drawBrush, 45.0F, 620.0F, drawFormat)
    '    e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 45.0F, 650.0F, drawFormat)

    '    e.Graphics.DrawString("...............................", drawFont, drawBrush, 400.0F, 510.0F, drawFormat)
    '    e.Graphics.DrawString("(Tandatangan Pegawai Pengesah)", drawFont, drawBrush, 400.0F, 530.0F, drawFormat)
    '    e.Graphics.DrawString("Nama:", drawFont, drawBrush, 400.0F, 560.0F, drawFormat)
    '    e.Graphics.DrawString("Jawatan:", drawFont, drawBrush, 400.0F, 590.0F, drawFormat)
    '    e.Graphics.DrawString("Jabatan:", drawFont, drawBrush, 400.0F, 620.0F, drawFormat)
    '    e.Graphics.DrawString("Tarikh:", drawFont, drawBrush, 400.0F, 650.0F, drawFormat)

    '    e.Graphics.DrawString("Nota:", drawFont, drawBrush, 755.0F, 440.0F, drawFormat)

    '    'horizon line
    '    e.Graphics.DrawLine(blackPen, 40.0F, 430.0F, 1120.0F, 430.0F)
    '    e.Graphics.DrawLine(blackPen, 40.0F, 680.0F, 1120.0F, 680.0F)

    '    'vertical line
    '    e.Graphics.DrawLine(blackPen, 40.0F, 430.0F, 40.0F, 680.0F)
    '    e.Graphics.DrawLine(blackPen, 395.0F, 430.0F, 395.0F, 680.0F)
    '    e.Graphics.DrawLine(blackPen, 750.0F, 430.0F, 750.0F, 680.0F)
    '    e.Graphics.DrawLine(blackPen, 1120.0F, 430.0F, 1120.0F, 680.0F)

    '    'detail
    '    e.Graphics.DrawString(datall.Rows(row).Item("NAMAPEMBEKAL"), drawFont1, drawbrush1, 45.0F, 130.0F, drawFormat)
    '    e.Graphics.DrawString(datall.Rows(row).Item("ALAMATPEMBEKAL"), drawFont1, drawbrush1, 45.0F, 150.0F, drawFormat)
    '    e.Graphics.DrawString(datall.Rows(row).Item("POSKOD"), drawFont1, drawbrush1, 45.0F, 170.0F, drawFormat)
    '    e.Graphics.DrawString(datall.Rows(row).Item("PESANAN"), drawFont1, drawbrush1, 870.0F, 130.0F, drawFormat)
    '    e.Graphics.DrawString(Label2.Text, drawFont1, drawbrush1, 870.0F, 190.0F, drawFormat)


    '    e.Graphics.DrawString(MetroTextBox3.Text, drawFont1, drawbrush1, 60.0F, 350.0F, drawFormat)
    '    e.Graphics.DrawString(MetroTextBox4.Text, drawFont1, drawbrush1, 130.0F, 350.0F, drawFormat)
    '    e.Graphics.DrawString("UNIT", drawFont1, drawbrush1, 400.0F, 350.0F, drawFormat)
    '    e.Graphics.DrawString(datall.Rows(row).Item("BILPESAN"), drawFont1, drawbrush1, 500.0F, 350.0F, drawFormat)
    '    e.Graphics.DrawString(datall.Rows(row).Item("CATATAN"), drawFont1, drawbrush1, 895.0F, 350.0F, drawFormat)

    '    If dahade Then
    '        e.Graphics.DrawString(g.ToString("N2"), drawFont1, drawbrush1, 685.0F, 350.0F, drawFormat)
    '        e.Graphics.DrawString(b.ToString, drawFont1, drawbrush1, 595.0F, 350.0F, drawFormat)

    '    Else
    '        e.Graphics.DrawString(d.ToString("N2"), drawFont1, drawbrush1, 685.0F, 350.0F, drawFormat)
    '        e.Graphics.DrawString(f.ToString, drawFont1, drawbrush1, 595.0F, 350.0F, drawFormat)

    '    End If
    '    e.Graphics.DrawString(totalsum.ToString("N2"), drawFont1, drawbrush1, 785.0F, 350.0F, drawFormat)
    'End Sub
    Private Sub printout()
        PrintPreviewDialog1.Document = PrintDocument3 'PrintPreviewDialog associate with PrintDocument.

        DirectCast(DirectCast(PrintPreviewDialog1.Controls(1), ToolStrip).Items(0), ToolStripButton).Enabled = False
        PrintPreviewDialog1.ShowDialog()

        PrintDialog1.Document = PrintDocument3 'PrintDialog associate with PrintDocument.

        If PrintDialog1.ShowDialog() = DialogResult.OK Then

            PrintDocument3.Print()

        End If
    End Sub

    'Private Sub PrintDocument3_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument3.PrintPage
    '    Dim drawFont As New Font("Arial", 12)
    '    Dim drawFont1 As New Font("Arial", 10)
    '    Dim drawFont2 As New Font("Arial", 14)
    '    Dim drawBrush As New SolidBrush(Color.DarkBlue)
    '    Dim drawbrush1 As New SolidBrush(Color.Black)
    '    Dim drawFormat As New StringFormat
    '    Dim blackPen As New Pen(Color.Black, 3)

    '    e.Graphics.DrawString("BORANG PENERIMAAN BARANG (BTB)", drawFont, drawBrush, 230.0F, 25.0F, drawFormat)
    '    e.Graphics.DrawString("(Tatacara Pengurusan Stor 43)", drawFont1, drawBrush, 300.0F, 45.0F, drawFormat)

    '    'penerangan
    '    e.Graphics.DrawString("Pegawai Pelulus", drawFont, drawBrush, 450.0F, 133.0F, drawFormat)
    '    e.Graphics.DrawString("Pengeluaran Stok", drawFont2, drawBrush, 60.0F, 145.0F, drawFormat)
    '    e.Graphics.DrawString("Kuantiti", drawFont2, drawBrush, 310.0F, 145.0F, drawFormat)
    '    e.Graphics.DrawString("  Kuantiti " + Environment.NewLine + "Diluluskan", drawFont, drawBrush, 420.0F, 160.0F, drawFormat)
    '    e.Graphics.DrawString("  Baki " + Environment.NewLine + "Kuantiti", drawFont, drawBrush, 540.0F, 160.0F, drawFormat)
    '    e.Graphics.DrawString("Catatan", drawFont2, drawBrush, 660.0F, 145.0F, drawFormat)

    '    e.Graphics.DrawString("..........................", drawFont2, drawBrush, 60.0F, 380.0F, drawFormat)
    '    e.Graphics.DrawString("(Tandatangan Pemohon)", drawFont2, drawBrush, 60.0F, 400.0F, drawFormat)
    '    e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 60.0F, 420.0F, drawFormat)
    '    e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 60.0F, 440.0F, drawFormat)
    '    e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 60.0F, 460.0F, drawFormat)

    '    e.Graphics.DrawString("..........................", drawFont2, drawBrush, 410.0F, 380.0F, drawFormat)
    '    e.Graphics.DrawString("(Tandatangan Pegawai Pelulus)", drawFont2, drawBrush, 410.0F, 400.0F, drawFormat)
    '    e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 410.0F, 420.0F, drawFormat)
    '    e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 410.0F, 440.0F, drawFormat)
    '    e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 410.0F, 460.0F, drawFormat)

    '    'horizon line
    '    e.Graphics.DrawLine(blackPen, 50.0F, 130.0F, 790.0F, 130.0F)
    '    e.Graphics.DrawLine(blackPen, 410.0F, 155.0F, 635.0F, 155.0F)
    '    e.Graphics.DrawLine(blackPen, 50.0F, 205.0F, 790.0F, 205.0F)
    '    e.Graphics.DrawLine(blackPen, 50.0F, 280.0F, 790.0F, 280.0F)
    '    e.Graphics.DrawLine(blackPen, 50.0F, 490.0F, 790.0F, 490.0F)
    '    'e.Graphics.DrawLine(blackPen, 120.0F, 250.0F, 750.0F, 250.0F)
    '    'e.Graphics.DrawLine(blackPen, 120.0F, 280.0F, 750.0F, 280.0F)

    '    'vertical line
    '    e.Graphics.DrawLine(blackPen, 50.0F, 130.0F, 50.0F, 490.0F)
    '    e.Graphics.DrawLine(blackPen, 300.0F, 130.0F, 300.0F, 280.0F)
    '    e.Graphics.DrawLine(blackPen, 410.0F, 130.0F, 410.0F, 490.0F)
    '    e.Graphics.DrawLine(blackPen, 525.0F, 155.0F, 525.0F, 280.0F)
    '    e.Graphics.DrawLine(blackPen, 635.0F, 130.0F, 635.0F, 280.0F)
    '    e.Graphics.DrawLine(blackPen, 790.0F, 130.0F, 790.0F, 490.0F)

    '    'detail
    '    e.Graphics.DrawString(MetroTextBox11.Text + " (" + MetroTextBox12.Text + ")", drawFont, drawbrush1, 60.0F, 215.0F, drawFormat)
    '    e.Graphics.DrawString(MetroTextBox15.Text, drawFont, drawbrush1, 305.0F, 215.0F, drawFormat)
    '    e.Graphics.DrawString(MetroTextBox17.Text, drawFont, drawbrush1, 415.0F, 215.0F, drawFormat)
    '    e.Graphics.DrawString(minus, drawFont, drawbrush1, 530.0F, 215.0F, drawFormat)
    '    ' e.Graphics.DrawString("eeeeeeeeeeeee", drawFont, drawbrush1, 640.0F, 215.0F, drawFormat)
    '    '-------------------------------------------------------------------------------------------------------------------------
    '    'penerangan
    '    e.Graphics.DrawString("Kemaskini Rekod:", drawFont2, drawBrush, 55.0F, 525.0F, drawFormat)
    '    e.Graphics.DrawString("Stok telah dikeluarkan dan" + Environment.NewLine + "direkod di Kad Petak No............", drawFont, drawBrush, 55.0F, 545.0F, drawFormat)
    '    e.Graphics.DrawString(".......................", drawFont2, drawBrush, 55.0F, 650.0F, drawFormat)
    '    e.Graphics.DrawString("(Tandatangan Pegawai Stor)", drawFont2, drawBrush, 55.0F, 670.0F, drawFormat)
    '    e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 55.0F, 690.0F, drawFormat)
    '    e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 55.0F, 710.0F, drawFormat)
    '    e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 55.0F, 730.0F, drawFormat)

    '    e.Graphics.DrawString("Perakuan Penerimaan:", drawFont2, drawBrush, 415.0F, 525.0F, drawFormat)
    '    e.Graphics.DrawString("Disahkan bahawa stok yang diluluskan" + Environment.NewLine + "telah diterima.", drawFont, drawBrush, 415.0F, 545.0F, drawFormat)
    '    e.Graphics.DrawString(".......................", drawFont2, drawBrush, 415.0F, 650.0F, drawFormat)
    '    e.Graphics.DrawString("(Tandatangan Pemohon)", drawFont2, drawBrush, 415.0F, 670.0F, drawFormat)
    '    e.Graphics.DrawString("Nama:", drawFont2, drawBrush, 415.0F, 690.0F, drawFormat)
    '    e.Graphics.DrawString("Jawatan:", drawFont2, drawBrush, 415.0F, 710.0F, drawFormat)
    '    e.Graphics.DrawString("Tarikh:", drawFont2, drawBrush, 415.0F, 730.0F, drawFormat)

    '    'horizontal line
    '    e.Graphics.DrawLine(blackPen, 50.0F, 515.0F, 790.0F, 515.0F)
    '    e.Graphics.DrawLine(blackPen, 50.0F, 760.0F, 790.0F, 760.0F)

    '    'vertical line
    '    e.Graphics.DrawLine(blackPen, 50.0F, 515.0F, 50.0F, 760.0F)
    '    e.Graphics.DrawLine(blackPen, 410.0F, 515.0F, 410.0F, 760.0F)
    '    e.Graphics.DrawLine(blackPen, 790.0F, 515.0F, 790.0F, 760.0F)
    'End Sub



    Private Sub MetroButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton2.Click
        Form2.ShowDialog()
    End Sub

    Private Sub MetroButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
       
    End Sub

    Private Sub MetroButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
      
    End Sub

    'Private Sub MetroButton11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton11.Click
    '    'printin()
    '    'printout()
    '    Dim thisDate As Date
    '    Dim thisMonth As Integer
    '    thisDate = Label2.Text
    '    thisMonth = Month(thisDate)
    '    MsgBox(thisMonth.ToString)
    'End Sub

    Private Sub MetroButton12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
      
    End Sub

    Private Sub ListView3_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
    Public Sub stok(ByVal lvw As ListView)
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        oleConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT * FROM kwlnstok", oleConn)
        da.Fill(ds, "stok")
        oleConn.Close()
        lvw.Clear()
        With lvw
            .Items.Clear()
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .Columns.Add("ID", 0, HorizontalAlignment.Left)
            .Columns.Add("KOD", 100, HorizontalAlignment.Left)
            .Columns.Add("NAMA ITEM", 230, HorizontalAlignment.Left)
            .Columns.Add("LOKASI", 100, HorizontalAlignment.Left)
            .Columns.Add("UNIT", 80, HorizontalAlignment.Left)
            .Columns.Add("KUMPULAN", 80, HorizontalAlignment.Left)
            .Columns.Add("TAHUN", 130, HorizontalAlignment.Left)
        End With
        For Each row As DataRow In ds.Tables("stok").Rows
            Dim lst As ListViewItem
            lst = lvw.Items.Add(row(0))
            For i As Integer = 1 To ds.Tables("stok").Columns.Count - 1
                lst.SubItems.Add(row(i))
            Next
        Next
    End Sub
    Public Sub kedudukan(ByVal lvw As ListView)
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        oleConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT * FROM stok", oleConn)
        da.Fill(ds, "ked")
        oleConn.Close()
        lvw.Clear()
        With lvw
            .Items.Clear()
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .Columns.Add("ID", 50, HorizontalAlignment.Left)
            .Columns.Add("KOD", 170, HorizontalAlignment.Left)
            '.Columns.Add("NAMA ITEM", 230, HorizontalAlignment.Left)
            '.Columns.Add("LOKASI", 100, HorizontalAlignment.Left)
            '.Columns.Add("UNIT", 80, HorizontalAlignment.Left)
            '.Columns.Add("KUMPULAN", 80, HorizontalAlignment.Left)
            '.Columns.Add("TAHUN", 130, HorizontalAlignment.Left)
        End With
        For Each row As DataRow In ds.Tables("ked").Rows
            Dim lst As ListViewItem
            lst = lvw.Items.Add(row(0))
            For i As Integer = 1 To ds.Tables("ked").Columns.Count - 1
                lst.SubItems.Add(row(i))
            Next
        Next
    End Sub


    'Private Sub MetroButton7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If MetroTextBox12.Text = "" Then
    '        MetroTextBox11.Text = ""
    '        MetroTextBox10.Text = ""
    '        MetroTextBox9.Text = ""
    '        MetroTextBox8.Text = ""
    '    Else

    '        Dim oleConn As System.Data.OleDb.OleDbConnection
    '        Dim da As OleDb.OleDbDataAdapter
    '        ds = New DataSet
    '        oleConn = New System.Data.OleDb.OleDbConnection
    '        oleConn.ConnectionString = My.Settings.manageConnectionString
    '        Try
    '            oleConn.Open()
    '            da = New OleDb.OleDbDataAdapter("select nama,harga,kuantiti,tarikh from detail where kod like '%" & MetroTextBox12.Text & "%'", oleConn)
    '            da.Fill(ds, "infox")
    '            oleConn.Close()

    '            MetroTextBox11.Text = ds.Tables("infox").Rows(row)("nama")
    '            MetroTextBox10.Text = ds.Tables("infox").Rows(row)("harga")
    '            MetroTextBox9.Text = ds.Tables("infox").Rows(row)("kuantiti")
    '            MetroTextBox8.Text = ds.Tables("infox").Rows(row)("tarikh")
    '            MetroLabel16.Visible = True
    '            MetroTextBox15.Visible = True
    '            dahade = True
    '        Catch ex As Exception
    '            MsgBox("Item Tiada Dalam Simpanan. Sila Masukkan Item Mengikut Seksyen Item Masuk Dahulu.")
    '            MetroTextBox11.Text = ""
    '            MetroTextBox10.Text = ""
    '            MetroTextBox9.Text = ""
    '            MetroTextBox8.Text = ""
    '            ' MetroLabel16.Visible = False
    '            'MetroTextBox15.Visible = False
    '            MetroTextBox12.Text = ""
    '            MetroTextBox12.Focus()
    '            dahade = False
    '        End Try
    '    End If

    'End Sub

    Private Sub MetroButton12_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton12.Click
        'Form4.row = row
        'Form4.ShowDialog()
        Form6.ShowDialog()
    End Sub

    'Private Sub MetroButton3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton3.Click
    '    If MetroTextBox4.Text = "" Or MetroTextBox5.Text = "" Or MetroTextBox6.Text = "" Or MetroTextBox7.Text = "" Then
    '        MsgBox("Sila Isi Ruangan Yang Kosong.")
    '    Else
    '        If dahade Then

    '            a = Convert.ToDecimal(MetroTextBox6.Text)
    '            b = Convert.ToDecimal(MetroTextBox14.Text)
    '            g = Convert.ToDecimal(MetroTextBox5.Text)
    '            'g1 = Format(MetroTextBox5.Text, "0.00")
    '            totalsum = (a + b) * g
    '            sum = a + b

    '            'Dim oleConn As System.Data.OleDb.OleDbConnection
    '            'Dim da As OleDb.OleDbDataAdapter
    '            'Dim add As String = ""
    '            'ds = New DataSet
    '            'oleConn = New System.Data.OleDb.OleDbConnection
    '            'oleConn.ConnectionString = My.Settings.manageConnectionString
    '            'oleConn.Open()
    '            'da = New OleDb.OleDbDataAdapter("UPDATE detail SET nama = '" & MetroTextBox4.Text & "', harga = '" & g.ToString("N2") & "', kuantiti = '" & sum.ToString() & "', tarikh = '" & Label2.Text & "', jumlah = '" & totalsum.ToString("N2") & "', namapemborong = '" & datall.Rows(row).Item("NAMAPEMBEKAL") & "', alamatpemborong = '" & datall.Rows(row).Item("ALAMATPEMBEKAL") & "', poskod = '" & datall.Rows(row).Item("POSKOD") & "', pesanan = '" & datall.Rows(row).Item("PESANAN") & "' WHERE kod LIKE '%" & MetroTextBox3.Text & "%'", oleConn)
    '            'da.Fill(ds, "infox")
    '            'oleConn.Close()

    '            Dim acsconn As System.Data.OleDb.OleDbConnection

    '            acsconn = New System.Data.OleDb.OleDbConnection
    '            acsconn.ConnectionString = My.Settings.manageConnectionString
    '            Dim cmd As New OleDb.OleDbCommand
    '            Try
    '                ',nilaikeluartahunan@cc,seksyen=@dd,baris=@ee,rak=@ff,tingkat=@gg,petak=@hh
    '                printin()
    '                acsconn.Open()
    '                cmd.Connection = acsconn
    '                cmd.CommandText = "update detail set kod=@a,nama=@b,harga=@c,kuantiti=@e,jumlah=@f,tarikh=@g,namapembekal=@h,alamatpembekal=@i,poskod=@j,pesanan=@k where kod = @a;"
    '                cmd.Parameters.AddWithValue("@a", MetroTextBox4.Text)
    '                cmd.Parameters.AddWithValue("@b", MetroTextBox4.Text)
    '                cmd.Parameters.AddWithValue("@c", g.ToString("N2"))
    '                'cmd.Parameters.AddWithValue("@d", MetroTextBox5.Text)
    '                cmd.Parameters.AddWithValue("@e", sum.ToString())
    '                cmd.Parameters.AddWithValue("@f", totalsum.ToString("N2"))
    '                cmd.Parameters.AddWithValue("@g", Label2.Text)
    '                cmd.Parameters.AddWithValue("@h", datall.Rows(row).Item("NAMAPEMBEKAL"))
    '                cmd.Parameters.AddWithValue("@i", datall.Rows(row).Item("ALAMATPEMBEKAL"))
    '                cmd.Parameters.AddWithValue("@j", datall.Rows(row).Item("POSKOD"))
    '                cmd.Parameters.AddWithValue("@k", datall.Rows(row).Item("PESANAN"))
    '                cmd.ExecuteNonQuery()
    '                acsconn.Close()
    '                readdatashow(ListView1)
    '                updateprice()
    '                MetroTextBox3.Text = ""
    '                MetroTextBox4.Text = ""
    '                MetroTextBox5.Text = ""
    '                MetroTextBox6.Text = ""
    '                MetroTextBox7.Text = ""
    '                MetroTextBox3.Focus()
    '                MsgBox("berjaya lagi")

    '            Catch ex As Exception
    '                MsgBox("gagal ubah")
    '            End Try

    '        Else

    '            d = Convert.ToDecimal(MetroTextBox5.Text)
    '            f = Convert.ToDecimal(MetroTextBox6.Text)
    '            totalsum = d.ToString("N2") * f.ToString("N2")
    '            printin()
    '            'Dim oleConn As System.Data.OleDb.OleDbConnection

    '            'oleConn = New System.Data.OleDb.OleDbConnection
    '            'oleConn.ConnectionString = My.Settings.manageConnectionString
    '            'Dim cmd As New OleDb.OleDbCommand
    '            ''Try
    '            'oleConn.Open()
    '            'cmd.Connection = oleConn
    '            'cmd.CommandText = "insert into detail(kod,nama,harga,kuantiti,tarikh,jumlah,namapemborong,alamatpemborong,poskod,pesanan) values('" + MetroTextBox3.Text + "','" + MetroTextBox4.Text + "','" + d.ToString("N2") + "','" + MetroTextBox6.Text + "','" + MetroTextBox7.Text + "','" + totalsum.ToString("N2") + "','" + datall.Rows(row).Item("NAMAPEMBEKAL") + "','" + datall.Rows(row).Item("ALAMATPEMBEKAL") + "','" + datall.Rows(row).Item("POSKOD") + "','" + datall.Rows(row).Item("PESANAN") + "')"
    '            'cmd.ExecuteNonQuery()
    '            '',namapemborong,alamatpemborong,poskod,pesanan
    '            '' "','" + datall.Rows(row).Item("NAMAPEMBEKAL") + "','" + datall.Rows(row).Item("ALAMATPEMBEKAL") + "','" + datall.Rows(row).Item("POSKOD") + "','" + datall.Rows(row).Item("PESANAN") + 
    '            'oleConn.Close()
    '            Dim acsconn As System.Data.OleDb.OleDbConnection

    '            acsconn = New System.Data.OleDb.OleDbConnection
    '            acsconn.ConnectionString = My.Settings.manageConnectionString
    '            Dim cmd As New OleDb.OleDbCommand
    '            Try
    '                ',nilaikeluartahunan@cc,seksyen=@dd,baris=@ee,rak=@ff,tingkat=@gg,petak=@hh
    '                printin()
    '                acsconn.Open()
    '                cmd.Connection = acsconn
    '                cmd.CommandText = "insert into detail(kod,nama,harga,kuantiti,tarikh,jumlah,namapembekal,alamatpembekal,poskod,pesanan) values(@a,@b,@c,@d,@e,@f,@g,@h,@i,@j);"
    '                cmd.Parameters.AddWithValue("@a", MetroTextBox3.Text)
    '                cmd.Parameters.AddWithValue("@b", MetroTextBox4.Text)
    '                cmd.Parameters.AddWithValue("@c", d.ToString("N2"))
    '                cmd.Parameters.AddWithValue("@d", MetroTextBox6.Text)
    '                cmd.Parameters.AddWithValue("@e", MetroTextBox7.Text)
    '                cmd.Parameters.AddWithValue("@f", totalsum.ToString("N2"))
    '                cmd.Parameters.AddWithValue("@g", datall.Rows(row).Item("NAMAPEMBEKAL"))
    '                cmd.Parameters.AddWithValue("@h", datall.Rows(row).Item("ALAMATPEMBEKAL"))
    '                cmd.Parameters.AddWithValue("@i", datall.Rows(row).Item("POSKOD"))
    '                cmd.Parameters.AddWithValue("@j", datall.Rows(row).Item("PESANAN"))
    '                'cmd.Parameters.AddWithValue("@g", id)
    '                cmd.ExecuteNonQuery()
    '                acsconn.Close()
    '                MsgBox("Pendaftaran Item Berjaya. Data Telah Disimpan.")
    '                readdatashow(ListView1)
    '                updateprice()
    '                MetroTextBox3.Text = ""
    '                MetroTextBox4.Text = ""
    '                MetroTextBox5.Text = ""
    '                MetroTextBox6.Text = ""
    '                MetroTextBox7.Text = ""
    '                MetroTextBox3.Focus()

    '            Catch ex As Exception
    '                MsgBox(ErrorToString)
    '            End Try
    '        End If

    '        MetroLabel15.Visible = False
    '        MetroTextBox14.Visible = False

    '    End If
    'End Sub

    Private Sub MetroButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    'Private Sub MetroButton4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If MetroTextBox3.Text = "" Then
    '        MetroTextBox4.Text = ""
    '        MetroTextBox5.Text = ""
    '        MetroTextBox6.Text = ""
    '        MetroTextBox7.Text = ""

    '    Else

    '        Dim oleConn As System.Data.OleDb.OleDbConnection

    '        Dim da As OleDb.OleDbDataAdapter
    '        ds = New DataSet
    '        oleConn = New System.Data.OleDb.OleDbConnection
    '        oleConn.ConnectionString = My.Settings.manageConnectionString
    '        Try
    '            oleConn.Open()
    '            da = New OleDb.OleDbDataAdapter("select kod,nama,harga,kuantiti,tarikh,jumlah,namapembekal,alamatpembekal,poskod,negeri,pesanan,kuantitipesan,kuantititerima from detail where kod like '%" & MetroTextBox3.Text & "%'", oleConn)
    '            da.Fill(ds, "infox")
    '            oleConn.Close()

    '            MetroTextBox4.Text = ds.Tables("infox").Rows(row)("nama")
    '            MetroTextBox5.Text = ds.Tables("infox").Rows(row)("harga")
    '            MetroTextBox6.Text = ds.Tables("infox").Rows(row)("kuantiti")
    '            MetroTextBox7.Text = ds.Tables("infox").Rows(row)("tarikh")

    '            MetroLabel15.Visible = True
    '            MetroTextBox14.Visible = True
    '            dahade = True
    '        Catch ex As Exception
    '            MsgBox("Item Tiada Dalam Simpanan. Sila Isi Setiap Perincian Item Pada Butang Tambah.")
    '            MetroTextBox4.Text = ""
    '            MetroTextBox5.Text = ""
    '            MetroTextBox6.Text = ""
    '            MetroTextBox7.Text = ""
    '            ' MetroTextBox7.Text = Label2.Text
    '            MetroLabel15.Visible = False
    '            MetroTextBox14.Visible = False
    '            dahade = False
    '        End Try
    '    End If

    ' End Sub

    Private Sub MetroButton9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton9.Click
        'Used printPreviewDialog instead of printdialog
        'to save paper while debugging.  When finished
        'swith out to PrintDialog.
        Dim PrintPreview As New PrintPreviewDialog
        PrintPreview.Document = PrintDocument1
        PrintPreview.ShowDialog()
    End Sub

    Private Sub MetroButton13_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton13.Click
        Form7.ShowDialog()
    End Sub

    Private Sub ListView4_DoubleClick1(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView4.DoubleClick
        Form5.row = ListView4.SelectedItems(row).Index
        Form5.Show()
    End Sub

    'Private Sub MetroButton8_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton8.Click
    '    'Dim a As Double = MetroTextBox9.Text
    '    'Dim b As Double = MetroTextBox15.Text
    '    'Dim minus As Double
    '    If MetroTextBox12.Text = "" Or MetroTextBox11.Text = "" Or MetroTextBox10.Text = "" Or MetroTextBox9.Text = "" Or MetroTextBox8.Text = "" Then
    '        MsgBox("Sila Isi Ruangan Yang Kosong")
    '    Else
    '        If MetroTextBox9.Text = "0" Then
    '            MsgBox("Tiada Item Di Dalam Stor.")
    '            MetroTextBox15.Text = ""
    '            MetroTextBox12.Text = ""
    '            MetroTextBox11.Text = ""
    '            MetroTextBox10.Text = ""
    '            MetroTextBox9.Text = ""
    '            MetroTextBox8.Text = ""
    '            MetroTextBox12.Focus()

    '        ElseIf b > a Then
    '            MsgBox("Bilangan Tolak Kuantiti Lebih Daripada Bilangan Asal. Sila Betulkan Bilangan Tolak Kuantiti.")
    '        Else
    '            printout()
    '            a = Convert.ToDecimal(MetroTextBox9.Text)
    '            b = Convert.ToDecimal(MetroTextBox15.Text)
    '            h = Convert.ToDecimal(MetroTextBox10.Text)
    '            minus = a - b
    '            totalsum = (a - b) * h
    '            Dim oleConn As System.Data.OleDb.OleDbConnection
    '            Dim da As OleDb.OleDbDataAdapter
    '            Dim add As String = ""
    '            ds = New DataSet
    '            oleConn = New System.Data.OleDb.OleDbConnection
    '            oleConn.ConnectionString = My.Settings.manageConnectionString
    '            oleConn.Open()
    '            da = New OleDb.OleDbDataAdapter("UPDATE detail SET nama = '" & MetroTextBox11.Text & "', harga = '" & h.ToString("N2") & "', kuantiti = '" & minus & "', tarikh = '" & Label2.Text & "', jumlah = '" & totalsum.ToString("N2") & "' WHERE kod LIKE '%" & MetroTextBox12.Text & "%'", oleConn)
    '            da.Fill(ds, "infox")
    '            oleConn.Close()
    '            readdatashow(ListView2)
    '            updateprice()

    '            MsgBox("Item Keluar Berjaya. Data Telah Disimpan.")
    '            'printout()
    '            MetroTextBox15.Text = ""
    '            MetroTextBox12.Text = ""
    '            MetroTextBox11.Text = ""
    '            MetroTextBox10.Text = ""
    '            MetroTextBox9.Text = ""
    '            MetroTextBox8.Text = ""
    '            MetroTextBox12.Focus()
    '        End If
    '    End If
    'End Sub


    Private Sub MetroButton14_Click(sender As Object, e As EventArgs) Handles MetroButton14.Click
        Form8.ShowDialog()
    End Sub


    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        Form9.ShowDialog()
    End Sub

    Private Sub ListView5_DoubleClick(sender As Object, e As EventArgs) Handles ListView5.DoubleClick
        Form12.row = ListView5.SelectedItems(row).Index
        Form12.Show()
    End Sub
End Class

