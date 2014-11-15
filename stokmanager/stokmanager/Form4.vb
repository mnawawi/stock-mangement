Public Class Form4
    Public Shared row As Integer
    Private Sub MetroButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton1.Click
        Me.Close()
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MetroTextBox13.Text = Form1.datall.Rows(row).Item("NAMAPEMBEKAL")
        MetroTextBox1.Text = Form1.datall.Rows(row).Item("ALAMATPEMBEKAL")
        MetroTextBox2.Text = Form1.datall.Rows(row).Item("POSKOD")
        MetroTextBox3.Text = Form1.datall.Rows(row).Item("PESANAN")
        MetroTextBox4.Text = Form1.datall.Rows(row).Item("CATATAN")
        MetroTextBox5.Text = Form1.datall.Rows(row).Item("BILPESAN")
    End Sub

    Private Sub MetroButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MetroButton3.Click
        Form1.datall.Rows(row).Item("NAMAPEMBEKAL") = MetroTextBox13.Text
        Form1.datall.Rows(row).Item("ALAMATPEMBEKAL") = MetroTextBox1.Text
        Form1.datall.Rows(row).Item("POSKOD") = MetroTextBox2.Text
        Form1.datall.Rows(row).Item("PESANAN") = MetroTextBox3.Text
        Form1.datall.Rows(row).Item("CATATAN") = MetroTextBox4.Text
        Form1.datall.Rows(row).Item("BILPESAN") = MetroTextBox5.Text
        Form1.detail = True
        Me.Close()
    End Sub
End Class