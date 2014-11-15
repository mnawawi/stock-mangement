Public Class Form11
    Dim ds As New DataSet
    Dim row As Integer

    Public Sub terimaan()
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        oleConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT * FROM tambah WHERE kod LIKE '%" & Form5.MetroTextBox13.Text & "%'", oleConn)
        da.Fill(ds, "terimaan")
        oleConn.Close()
        ListView1.Clear()
        With ListView1
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
        For Each row As DataRow In ds.Tables("terimaan").Rows
            Dim lst As ListViewItem
            lst = ListView1.Items.Add(row(0))
            For i As Integer = 1 To ds.Tables("terimaan").Columns.Count - 1
                lst.SubItems.Add(row(i))
            Next
        Next
    End Sub
    Public Sub keluaran()
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        oleConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT * FROM kurang WHERE kod LIKE '%" & Form5.MetroTextBox13.Text & "%'", oleConn)
        da.Fill(ds, "keluaran")
        oleConn.Close()
        ListView2.Clear()
        With ListView2
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
        For Each row As DataRow In ds.Tables("keluaran").Rows
            Dim lst As ListViewItem
            lst = ListView2.Items.Add(row(0))
            For i As Integer = 1 To ds.Tables("keluaran").Columns.Count - 1
                lst.SubItems.Add(row(i))
            Next
        Next
    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MetroLabel3.Text = Form5.MetroTextBox13.Text
        MetroLabel5.Text = Form5.MetroLabel1.Text
        terimaan()
        keluaran()
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        Me.Close()
    End Sub
End Class