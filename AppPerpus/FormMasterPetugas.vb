Imports System.Data.SqlClient
Public Class FormMasterPetugas

    Sub KondisiAwal()

        ' Memeberishkan textbox setelah input data
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""

        Call Koneksi()
        Da = New SqlDataAdapter("select KodePetugas, NamaPetugas, LevelPetugas from TBL_PETUGAS", Conn)
        Ds = New DataSet
        Da.Fill(Ds, "TBL_PETUGAS")
        DataGridView1.DataSource = (Ds.Tables("TBL_PETUGAS"))
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("ADMIN")
        ComboBox1.Items.Add("USER")

        TextBox3.PasswordChar = "*"

    End Sub

    Private Sub FormMasterPetugas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        ' Validasi Data untuk membatasi panjang string yang ada pada textbox sesuai database
        TextBox1.MaxLength = 6
        TextBox2.MaxLength = 50
        TextBox3.MaxLength = 30
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Pastikan Semua Field Terisi")
        Else
            Call Koneksi()
            Dim SimpanData As String =
                "insert into TBL_PETUGAS values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "')"
            Cmd = New SqlCommand(SimpanData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Simpan")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Pastikan Semua Field Terisi")
        Else
            Call Koneksi()
            Dim EditData As String =
                "update TBL_PETUGAS set NamaPetugas='" & TextBox2.Text & "',PasswordPetugas='" & ComboBox1.Text & "',LevelPetugas='" & TextBox3.Text & "' where KodePetugas='" & TextBox1.Text & "'"
            Cmd = New SqlCommand(EditData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Edit")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or ComboBox1.Text = "" Then
            MsgBox("Pastikan Semua Field Terisi")
        Else
            Call Koneksi()
            Dim HapusData As String =
                "delete TBL_PETUGAS where KodePetugas='" & TextBox1.Text & "'"
            Cmd = New SqlCommand(HapusData, Conn)
            Cmd.ExecuteNonQuery()
            MsgBox("Data Berhasil Di Hapus")
            Call KondisiAwal()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New SqlCommand("select * from TBL_PETUGAS where KodePetugas='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Rd.HasRows Then
                TextBox2.Text = Rd.Item("NamaPetugas")
                TextBox3.Text = Rd.Item("PasswordPetugas")
                ComboBox1.Text = Rd.Item("LevelPetugas")
            Else
                MsgBox("Data Tidak Ada")
            End If
        End If
    End Sub
End Class