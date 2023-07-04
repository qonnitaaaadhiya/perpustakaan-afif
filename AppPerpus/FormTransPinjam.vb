Imports System.Data.SqlClient
Public Class FormTransPinjam
    Sub KondisiAwal()
        Call NoOtomatis()
        Call MunculKodeAnggota()
        LBLNamaPetugas.Text = FormMenuUtama.STLabel4.Text
        LBLNama.Text = ""
        LBLAlamat.Text = ""
        LBLTelepon.Text = ""
        LBLJudul.Text = ""
        LBLPengarang.Text = ""
        LBLTahun.Text = ""
        LBLTotalBuku.Text = "0"
        ComboBox1.Text = ""
        Call BuatKolom()
    End Sub

    Private Sub FormTransPinjam_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call KondisiAwal()
        LBLTanggal.Text = Today
    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LBLJam.Text = TimeOfDay
    End Sub
    Sub NoOtomatis()
        Call Koneksi()
        Cmd = New SqlCommand("select * from TBL_PINJAM where NoPinjam in (select max(NoPinjam) from TBL_PINJAM)", Conn)
        Dim UrutkanKode As String
        Dim Hitung As Long
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Not Rd.HasRows Then
            UrutkanKode = "PJM" + Format(Now, "yyMMdd") + "001"
        Else
            Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
            UrutkanKode = "PJM" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
        End If
        LBLNoPinjam.Text = UrutkanKode
        'TextBox2.Focus()
    End Sub
    Sub MunculKodeAnggota()
        Call Koneksi()
        ComboBox1.Items.Clear()
        Cmd = New SqlCommand("Select * from tbl_anggota", Conn)
        Rd = Cmd.ExecuteReader
        Do While Rd.Read
            ComboBox1.Items.Add(Rd.Item(0))
        Loop
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Call Koneksi()
        Cmd = New SqlCommand("select * from TBL_ANGGOTA where KodeAnggota ='" & ComboBox1.Text & "'", Conn)
        Rd = Cmd.ExecuteReader
        Rd.Read()
        If Rd.HasRows Then
            LBLNama.Text = Rd!NamaAnggota
            LBLAlamat.Text = Rd!AlamatAnggota
            LBLTelepon.Text = Rd!TelpAnggota
            Call CekPinjaman()
        End If
    End Sub
    Sub BuatKolom()
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Kode", "Kode Buku")
        DataGridView1.Columns.Add("Judul", "Judul Buku")
        DataGridView1.Columns.Add("Pengarang", "Pengarang")
        DataGridView1.Columns.Add("Penerbit", "Penerbit")
        DataGridView1.Columns.Add("Jumlah", "Jumlah")
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If e.KeyChar = Chr(13) Then
            Call Koneksi()
            Cmd = New SqlCommand("Select * From TBL_BUKU where KodeBuku='" & TextBox1.Text & "'", Conn)
            Rd = Cmd.ExecuteReader
            Rd.Read()
            If Not Rd.HasRows Then
                MsgBox("Kode Buku tidak ada")
            Else
                TextBox1.Text = Rd.Item("KodeBuku")
                LBLJudul.Text = Rd.Item("JudulBuku")
                LBLPengarang.Text = Rd.Item("PengarangBuku")
                LBLTahun.Text = Rd.Item("TahunBuku")
                TextBox2.Enabled = True
                TextBox2.Text = "1"
            End If
        End If
    End Sub
    Sub RumusTotalBuku()
        Dim HitungItem As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            HitungItem = HitungItem + DataGridView1.Rows(i).Cells(4).Value
            LBLTotalBuku.Text = HitungItem
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Val(LBLTelahPinjam.Text) + Val(LBLTotalBuku.Text) >= 5 Or Val(LBLTotalBuku.Text) + Val(TextBox2.Text) > 5 Then
            MsgBox("Buku yang di pinjam melebihi maksimal !")
        Else
            If LBLJudul.Text = "" Or TextBox2.Text = "" Then
                MsgBox("Silahkan masukan Kode Buku dan Tekan ENTER!")
            Else
                DataGridView1.Rows.Add(New String() {TextBox1.Text, LBLJudul.Text, LBLPengarang.Text, LBLTahun.Text, TextBox2.Text})
                TextBox1.Text = ""
                TextBox2.Text = ""
                LBLJudul.Text = ""
                TextBox2.Text = ""
                LBLPengarang.Text = ""
                LBLTahun.Text = ""
                Call RumusTotalBuku()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If LBLNama.Text = "" Or Label9.Text = "" Then
            MsgBox("Transaksi Tidak Ada, silahkan lalukan transaksi terleih dahulu")
        Else
            Call Koneksi()
            Dim tglsql As String
            tglsql = Format(Today, "yyyy-MM-dd")
            Dim PinjamBuku As String = "Insert into TBL_PINJAM values ('" & LBLNoPinjam.Text & "', '" & tglsql & "', '" & LBLJam.Text & "', '" & ComboBox1.Text & "', '" & LBLTotalBuku.Text & "', '" & LBLTotalBuku.Text & "', '" & FormMenuUtama.STLabel2.Text & "')"
            Cmd = New SqlCommand(PinjamBuku, Conn)
            Cmd.ExecuteNonQuery()

            For Baris As Integer = 0 To DataGridView1.Rows.Count - 2
                Call Koneksi()
                Dim SimpanDetail As String = "Insert into TBL_DETAILPINJAM values('" & LBLNoPinjam.Text & "','" & DataGridView1.Rows(Baris).Cells(0).Value & "', '" & DataGridView1.Rows(Baris).Cells(4).Value & "')"
                Cmd = New SqlCommand(SimpanDetail, Conn)
                Cmd.ExecuteNonQuery()
                Call Koneksi()
                Cmd = New SqlCommand("select * from TBL_BUKU where KodeBuku='" & DataGridView1.Rows(Baris).Cells(0).Value & "'", Conn)
                Rd = Cmd.ExecuteReader
                Rd.Read()
                Call Koneksi()
                Dim Kurangistok As String = "Update TBL_BUKU set StokBuku='" & Rd.Item("StokBuku") - DataGridView1.Rows(Baris).Cells(4).Value & "'where KodeBuku='" & DataGridView1.Rows(Baris).Cells(0).Value & "'"
                Cmd = New SqlCommand(Kurangistok, Conn)
                Cmd.ExecuteNonQuery()
            Next
            Call KondisiAwal()
            MsgBox("Transaksi Telah Berhasil Disimpan")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call KondisiAwal()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Sub CekPinjaman()
        Call Koneksi()
        Da = New SqlDataAdapter("Select TBL_BUKU.KodeBuku, JudulBuku, JumlahBuku From TBL_ANGGOTA,TBL_PINJAM,TBL_BUKU,TBL_DETAILPINJAM where TBL_BUKU.KodeBuku=TBL_DETAILPINJAM.KodeBuku And TBL_PINJAM.NoPinjam=TBL_DETAILPINJAM.NoPinjam And TBL_ANGGOTA.KodeAnggota=TBL_PINJAM.KodeAnggota And TBL_Anggota.KodeAnggota='" & ComboBox1.Text & "' AND TBL_DETAILPINJAM.JumlahBuku>0", Conn)
        Ds = New DataSet
        Ds.Clear()
        Da.Fill(Ds, "Detail")
        DataGridView2.DataSource = Ds.Tables("Detail")
        LBLTelahPinjam.Text = DataGridView2.Rows.Count - 1
        DataGridView2.ReadOnly = True
    End Sub
End Class