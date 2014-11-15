Public Class Form2
    Public Shared ds As DataSet
    Public Shared row As Integer



    Public Sub readdatashow(ByVal lvw As ListView)

        Dim oleConn As System.Data.OleDb.OleDbConnection
        Dim da As OleDb.OleDbDataAdapter
        ds = New DataSet
        oleConn = New System.Data.OleDb.OleDbConnection
        oleConn.ConnectionString = My.Settings.manageConnectionString
        oleConn.Open()
        da = New OleDb.OleDbDataAdapter("SELECT * FROM staff", oleConn)
        da.Fill(ds, "infox")
        oleConn.Close()
        lvw.Clear()
        With lvw
            .Items.Clear()
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .Columns.Add("ID", 0, HorizontalAlignment.Left)
            .Columns.Add("NAMA PENGGUNA", 150, HorizontalAlignment.Left)
            .Columns.Add("KATA LALUAN", 150, HorizontalAlignment.Left)
            .Columns.Add("NAMA", 150, HorizontalAlignment.Left)
            .Columns.Add("NO. KAD PENGENALAN", 150, HorizontalAlignment.Left)
        End With
        For Each row As DataRow In ds.Tables("infox").Rows
            Dim lst As ListViewItem
            lst = lvw.Items.Add(row(0))
            For i As Integer = 1 To ds.Tables("infox").Columns.Count - 1
                lst.SubItems.Add(row(i))
            Next
        Next
        'With lvw
        '    .Items.Clear()
        '    .View = View.Details
        '    .GridLines = True
        '    .FullRowSelect = True
        '    .Columns.Add("ID", 0, HorizontalAlignment.Left)
        '    .Columns.Add("NAMA", 150, HorizontalAlignment.Left)
        '    .Columns.Add("NO. KP", 150, HorizontalAlignment.Left)
        '    .Columns.Add("NAMA PENGGUNA", 150, HorizontalAlignment.Left)
        '    .Columns.Add("KATA LALUAN", 150, HorizontalAlignment.Left)
        'End With





    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        readdatashow(ListView1)
    End Sub

    Private Sub MetroButton8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton8.Click
        ' MsgBox(ds.Tables("infox").Rows(row)("namastaf"))
        'If ds.Tables("infox").Rows.Count > 0 Then
        '    MsgBox(ds.Tables("infox").Rows(row)("namastaf"))
        'Else
        '    MsgBox("empty")
        'End If
        Dim OK As Boolean = False
        Dim namastaf As String = MetroTextBox12.Text
        Dim nokp As String = MetroTextBox11.Text
        Dim username As String = MetroTextBox9.Text
        Dim password As String = MetroTextBox8.Text
        If ds.Tables("infox").Rows.Count > 0 Then
            For i As Integer = 0 To ds.Tables("infox").Rows.Count - 1
                If MetroTextBox12.Text = ds.Tables("infox").Rows(i)("namastaf") Or MetroTextBox11.Text = ds.Tables("infox").Rows(i)("nokp") Then
                    OK = False
                    MsgBox("ID DIPILIH SUDAH ADA DALAM PENGKALAN DATA, SILA UBAH ID!")
                    Exit For
                Else
                    OK = True
                End If
            Next
            If OK Then
                Dim oleConn As System.Data.OleDb.OleDbConnection

                oleConn = New System.Data.OleDb.OleDbConnection
                oleConn.ConnectionString = My.Settings.manageConnectionString
                Dim cmd As New OleDb.OleDbCommand
                Try
                    oleConn.Open()
                    cmd.Connection = oleConn
                    cmd.CommandText = "INSERT INTO staff(username,katalaluan,namastaf,nokp,tarikh,masa) VALUES(@a,@b,@c,@d,@e,@f);"
                    cmd.Parameters.AddWithValue("@a", username)
                    cmd.Parameters.AddWithValue("@b", password)
                    cmd.Parameters.AddWithValue("@c", namastaf)
                    cmd.Parameters.AddWithValue("@d", nokp)
                    cmd.Parameters.AddWithValue("@e", Form1.Label2.Text)
                    cmd.Parameters.AddWithValue("@f", Form1.Label1.Text + " " + Form1.Label3.Text)
                    cmd.ExecuteNonQuery()

                    oleConn.Close()

                    MsgBox("ID DAN KATALALUAN TELAH DISIMPAN")
                    readdatashow(ListView1)
                    MetroTextBox12.Text = ""
                    MetroTextBox11.Text = ""
                    MetroTextBox9.Text = ""
                    MetroTextBox8.Text = ""
                Catch ex As Exception
                    MsgBox(ErrorToString)
                End Try
            End If
           
            ' MsgBox("empty")
        Else
            Dim oleConn1 As System.Data.OleDb.OleDbConnection

            oleConn1 = New System.Data.OleDb.OleDbConnection
            oleConn1.ConnectionString = My.Settings.manageConnectionString
            Dim cmd1 As New OleDb.OleDbCommand
            Try
                oleConn1.Open()
                cmd1.Connection = oleConn1
                cmd1.CommandText = "INSERT INTO staff(username,katalaluan,namastaf,nokp) VALUES(@a,@b,@c,@d);"
                cmd1.Parameters.AddWithValue("@a", username)
                cmd1.Parameters.AddWithValue("@b", password)
                cmd1.Parameters.AddWithValue("@c", namastaf)
                cmd1.Parameters.AddWithValue("@d", nokp)
                cmd1.ExecuteNonQuery()

                oleConn1.Close()

                MsgBox("ID DAN KATALALUAN TELAH DISIMPAN")
                readdatashow(ListView1)
                MetroTextBox12.Text = ""
                MetroTextBox11.Text = ""
                MetroTextBox9.Text = ""
                MetroTextBox8.Text = ""
            Catch ex As Exception
                MsgBox(ErrorToString)
            End Try
        End If






    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        ' Form3.row = ds.Tables("infox").Rows(row)("ID")
        Form3.row = ListView1.SelectedItems(row).Index
        'indx = ListView3.FocusedItem.Index
        Form3.ShowDialog()
    End Sub

    Private Sub MetroButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton7.Click
        Me.Close()
    End Sub
End Class

