Public Class Form3
    Dim id As String
    Public Shared row As Integer
    Dim ds As DataSet
    Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        id = Form2.ds.Tables("infox").Rows(row)("ID")
        MetroTextBox12.Text = Form2.ds.Tables("infox").Rows(row)("namastaf")
        MetroTextBox11.Text = Form2.ds.Tables("infox").Rows(row)("nokp")
        MetroTextBox9.Text = Form2.ds.Tables("infox").Rows(row)("username")
        MetroTextBox8.Text = Form2.ds.Tables("infox").Rows(row)("katalaluan")

    End Sub

    Private Sub MetroButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton8.Click
        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        Dim add As String = ""
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        oleConn.Open()
        da = New OleDb.OleDbDataAdapter("UPDATE staff SET username = '" & MetroTextBox9.Text & "', katalaluan = '" & MetroTextBox8.Text & "', namastaf = '" & MetroTextBox12.Text & "', nokp = '" & MetroTextBox11.Text & "' WHERE nokp LIKE '%" & MetroTextBox11.Text & "%'", oleConn)
        da.Fill(ds, "infox")
        oleConn.Close()

        Me.Close()
        Form2.readdatashow(Form2.ListView1)
    End Sub

    Private Sub MetroButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton7.Click
        Me.Close()
    End Sub

    Private Sub MetroButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton1.Click
        Dim oleConn As System.Data.OleDb.OleDbConnection
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        Dim cmd As New OleDb.OleDbCommand
        oleConn.Open()
        cmd.Connection = oleConn
        cmd.CommandText = "DELETE FROM staff WHERE ID = " + id
        cmd.ExecuteNonQuery()
        oleConn.Close()
        Me.Close()
        Form2.readdatashow(Form2.ListView1)
    End Sub
End Class