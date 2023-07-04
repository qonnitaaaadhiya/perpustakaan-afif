Imports System.Data.SqlClient
Public Class FormTransKembali
	Private Sub KondisiAwal()
		TxtKodeAgt.Text = ""
		LBLNamaAgt.Text = ""
		TextBox1.Text = ""
		TextBox2.Text = ""
		LBLJudul.Text = ""
		TextBox2.Text = ""
		LBLTglPjm.Text = ""
		LBLLamaPjm.Text = ""
		LBL_Denda.Text = ""
		LBLNoPjm.Text = ""
		TxtDiBayar.Text = ""
		LBLDenda.Text = ""
		LBLKembali.Text = ""
		LBLTelahPinjam.Text = ""
	End Sub
	Private Sub KondisiAwal2()
		TextBox1.Text = ""
		TextBox2.Text = ""
		LBLJudul.Text = ""
		TextBox2.Text = ""
		LBLTglPjm.Text = ""
		LBLLamaPjm.Text = ""
		LBL_Denda.Text = ""
		LBLNoPjm.Text = ""
		TxtDiBayar.Text = ""
		LBLDenda.Text = ""
		LBLKembali.Text = ""
	End Sub
	Sub TampilGrid1()
		DataGridView1.Columns.Clear()
		DataGridView1.Columns.Add("Kode", "KodeBuku")
		DataGridView1.Columns.Add("Nomor", "Nomor Pinjam")
		DataGridView1.Columns.Add("Judul", "Judul Buku")
		DataGridView1.Columns.Add("Jumlah", "Jumlah Pinjam")
		DataGridView1.Columns.Add("Tanggal", "Tanggal Pinjam")
		DataGridView1.Columns.Add("Lama", "Lama Pinjam")
		DataGridView1.Columns.Add("Denda", "Denda")
		DataGridView1.Columns(0).Width = 50
		DataGridView1.Columns(1).Width = 75
		DataGridView1.Columns(2).Width = 150
		DataGridView1.Columns(3).Width = 50
		DataGridView1.Columns(4).Width = 75
		DataGridView1.Columns(5).Width = 50
		DataGridView1.Columns(6).Width = 75
	End Sub
	Sub TampilGrid2()
		DataGridView2.Columns.Add("Kode", "Kode")
		DataGridView2.Columns.Add("Judul", "Judul Buku")
		DataGridView2.Columns.Add("Jumlah", "Jumlah")
		DataGridView2.Columns(0).Width = 75
		DataGridView2.Columns(1).Width = 300
		DataGridView2.Columns(2).Width = 100
		DataGridView2.ReadOnly = True
	End Sub
	Sub Otomatis()
		Call Koneksi()
		Cmd = New SqlCommand("select * from TBL_KEMBALI where NoKembali in (select max(NoKembali) from TBL_KEMBALI)", Conn)
		Dim UrutkanKode As String
		Dim Hitung As Long
		Rd = Cmd.ExecuteReader
		Rd.Read()
		If Not Rd.HasRows Then
			UrutkanKode = "KBL" + Format(Now, "yyMMdd") + "001"
		Else
			Hitung = Microsoft.VisualBasic.Right(Rd.GetString(0), 9) + 1
			UrutkanKode = "KBL" + Format(Now, "yyMMdd") + Microsoft.VisualBasic.Right("000" & Hitung, 3)
		End If
		LBLNomorKbl.Text = UrutkanKode
		'TextBox2.Focus()
	End Sub
	Private Sub FormTransKembali_Activated(sender As Object, e As EventArgs) Handles Me.Activated
		Koneksi()
		Call Otomatis()
		LBLTanggalKbl.Text = Today
	End Sub
	Private Sub FormTransKembali_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Call KondisiAwal()
		Call Koneksi()
		Call TampilGrid1()
		Call Otomatis()
		LBLNamaPetugas.Text = FormMenuUtama.STLabel4.Text
	End Sub
	Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
		LBLJam.Text = TimeOfDay
	End Sub
	Sub Pinjaman()
		Call Koneksi()
		Da = New SqlDataAdapter("Select TBL_BUKU.KodeBuku, JudulBuku, JumlahBuku From TBL_ANGGOTA, TBL_PINJAM, TBL_BUKU, TBL_DETAILPINJAM where TBL_BUKU.KodeBuku=TBL_DETAILPINJAM.KodeBuku and TBL_PINJAM.NoPinjam=TBL_DETAILPINJAM.NoPinjam and TBL_ANGGOTA.kodeAnggota = TBL_PINJAM.KodeAnggota and TBL_ANGGOTA.KodeAnggota = '" & TxtKodeAgt.Text & "' and TBL_DETAILPINJAM.JumlahBuku>0", Conn)
		Ds = New DataSet
		Ds.Clear()
		Da.Fill(Ds, "Detail")
		DataGridView2.DataSource = Ds.Tables("Detail")
		DataGridView2.ReadOnly = True
		DataGridView2.Columns(0).Width = 80
		DataGridView2.Columns(1).Width = 300
		DataGridView2.Columns(2).Width = 100
	End Sub
	Sub CariData()
		Call Koneksi()
		Cmd = New SqlCommand("Select * from TBL_ANGGOTA where KodeAnggota='" & TxtKodeAgt.Text & "'", Conn)
		Rd = Cmd.ExecuteReader
		Rd.Read()
		If Rd.HasRows Then
			LBLNamaAgt.Text = Rd.Item("NamaAnggota")
			Call Koneksi()
			Cmd = New SqlCommand("Select KodeAnggota From TBL_PINJAM where KodeAnggota ='" & TxtKodeAgt.Text & "'", Conn)
			Rd = Cmd.ExecuteReader
			Rd.Read()
			If Rd.HasRows Then
				Call Koneksi()
				Cmd = New SqlCommand("Select SUM(TotalPinjam) as Ketemu from TBL_PINJAM where KodeAnggota='" & TxtKodeAgt.Text & "'", Conn)
				Rd = Cmd.ExecuteReader
				Rd.Read()
				LBLTelahPinjam.Text = Rd.Item(0)
				Call Pinjaman()
			End If
		End If
	End Sub

	Private Sub TxtKodeAgt_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtKodeAgt.KeyPress
		If e.KeyChar = Chr(13) Then
			Call CariData()
		End If
	End Sub
	Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
		If e.KeyChar = Chr(13) Then
			Call Koneksi()
			Cmd = New SqlCommand("Select distinct TBL_BUKU.KodeBuku,TBL_DETAILPINJAM.NoPinjam,JudulBuku,JumlahBuku,TglPinjam From TBL_ANGGOTA,TBL_PINJAM,TBL_BUKU,TBL_DETAILPINJAM where TBL_BUKU.KodeBuku=TBL_DETAILPINJAM.KodeBuku and TBL_PINJAM.NoPinjam = TBL_DETAILPINJAM.NoPinjam and TBL_ANGGOTA.KodeAnggota = TBL_PINJAM.KodeAnggota and TBL_ANGGOTA.KodeAnggota = '" & TxtKodeAgt.Text & "' and TBL_DETAILPINJAM.KodeBuku = '" & TextBox1.Text & "' and TBL_DETAILPINJAM.JumlahBuku>0", Conn)
			Rd = Cmd.ExecuteReader
			Rd.Read()
			If Not Rd.HasRows Then
				MsgBox("Kode Buku Bukan Yang di Pinjam")
			Else
				LBLNoPjm.Text = Rd.Item(1)
				LBLJudul.Text = Rd.Item(2)
				LBLTglPjm.Text = Rd.Item(4)
				TextBox2.Text = Rd.Item(3)
				LBLLamaPjm.Text = DateDiff(DateInterval.Day, Rd.Item(4), Today())
				If LBLLamaPjm.Text > 5 Then
					LBL_Denda.Text = Val((LBLLamaPjm.Text) - 5) * 1000
				Else
					LBL_Denda.Text = 0
				End If
			End If
		End If
	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		DataGridView1.Rows.Add(New String() {TextBox1.Text, LBLNoPjm.Text, LBLJudul.Text, TextBox2.Text, LBLTglPjm.Text, LBLLamaPjm.Text, LBL_Denda.Text})
		Call HapusMaster()
		Call TotalDENDA()
		Call TotalKembali()
	End Sub
	Private Sub HapusMaster()
		TextBox1.Text = ""
		LBLNoPjm.Text = ""
		LBLJudul.Text = ""
		LBLTglPjm.Text = ""
		TextBox2.Text = ""
		LBLLamaPjm.Text = ""
		LBL_Denda.Text = ""
	End Sub
	Sub TotalDENDA()
		Dim Subtot As Integer = 0
		For I As Integer = 0 To DataGridView1.Rows.Count - 1
			Subtot += Val(DataGridView1.Rows(I).Cells(6).Value)
			LBLDenda.Text = Subtot
		Next
	End Sub
	Sub TotalKembali()
		Dim Subtot As Integer = 0
		For I As Integer = 0 To DataGridView1.Rows.Count - 1
			Subtot += Val(DataGridView1.Rows(I).Cells(3).Value)
			LBLTotal.Text = Subtot
		Next
	End Sub

	Private Sub TxtDiBayar_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDiBayar.KeyPress
		If e.KeyChar = Chr(13) Then
			If Val(TxtDiBayar.Text) < Val(LBLDenda.Text) Then
				MsgBox("Pembayaran Kurang")
				TxtDiBayar.Focus()
			ElseIf Val(TxtDiBayar.Text) = Val(LBLDenda.Text) Then
				LBLKembali.Text = 0
				Button2.Focus()
			Else
				LBLKembali.Text = Val(LBLDenda.Text) - Val(TxtDiBayar.Text)
				Button2.Focus()
			End If
		End If
	End Sub

	Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
		Me.Close()
	End Sub

	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		If TxtKodeAgt.Text = "" Or LBLDenda.Text = "" Or LBLKembali.Text = "" Then
			MsgBox("Transaksi Tidak ada, slahkan lakukan transaksi terlebih dahulu")
		Else
			Call Koneksi()
			Dim tglsql As String
			tglsql = Format(Today, "yyyy-MM-dd")
			Dim BukuKembali As String = "Insert into TBL_KEMBALI(NoKembali, TglKembali, TotalKembali, Denda, Dibayar, Kembali, KodeAnggota, KodePetugas) values ('" & LBLNomorKbl.Text & "', '" & tglsql & "', '" & LBLTotal.Text & "','" & TxtDiBayar.Text & "','" & LBLDenda.Text & "','" & LBLKembali.Text & "', '" & TxtKodeAgt.Text & "', '" & FormMenuUtama.STLabel2.Text & "')"
			Cmd = New SqlCommand(BukuKembali, Conn)
			Cmd.ExecuteNonQuery()

			For Baris As Integer = 0 To DataGridView1.Rows.Count = 2
				Call Koneksi()
				Dim SimpanDetail As String = "Insert into tbl_detailkembali(NoKembali,KodeBuku,JumlahBuku) values('" & LBLNomorKbl.Text & "', '" & DataGridView1.Rows(Baris).Cells(0).Value & "','" & DataGridView1.Rows(Baris).Cells(3).Value & "')"
				Cmd = New SqlCommand(SimpanDetail, Conn)
				Cmd.ExecuteNonQuery()

				'Update Buku
				Call Koneksi()
				Cmd = New SqlCommand("select * from tbl_buku where kodebuku='" & DataGridView1.Rows(Baris).Cells(0).Value & "'", Conn)
				Rd = Cmd.ExecuteReader
				Rd.Read()
				Call Koneksi()
				Dim Kurangstok As String = "Update tbl_buku set StokBuku= '" & Rd.Item("StokBuku") + DataGridView1.Rows(Baris).Cells(3).Value & "'where KodeBuku='" & DataGridView1.Rows(Baris).Cells(0).Value & "'"
				Cmd = New SqlCommand(Kurangstok, Conn)
				Cmd.ExecuteNonQuery()

				'Update Detail Pinjam
				Call Koneksi()
				Cmd = New SqlCommand("select JumlahBuku from TBL_DETAILPINJAM where KodeBuku='" & DataGridView1.Rows(Baris).Cells(0).Value & "' and NoPinjam='" & DataGridView1.Rows(Baris).Cells(1).Value & "'", Conn)

				Rd.Read()
				If Rd.HasRows Then
					Call Koneksi()
					Dim updatedetailpinjam As String = "update TBL_DETAILPINJAM set JumlahBuku= '" & Rd.Item(0) - DataGridView1.Rows(Baris).Cells(3).Value & "'where KodeBuku='" & DataGridView1.Rows(Baris).Cells(0).Value & "' and NoPinjam='" & DataGridView1.Rows(Baris).Cells(1).Value & "'"
					Cmd = New SqlCommand(updatedetailpinjam, Conn)
					Cmd.ExecuteNonQuery()
				End If
				 
				'Update Total Pinjam
				Call Koneksi()
				Cmd = New SqlCommand("select totalPinjam from TBL_pinjam where NoPinjam='" & DataGridView1.Rows(Baris).Cells(1).Value & "'", Conn)
				Rd = Cmd.ExecuteReader
				Rd.Read()
				If Rd.HasRows Then
					Call Koneksi()
					Dim Sqleditpjm As String = "update TBL_pinjam set totalPinjam= '" & Rd.Item(0) - DataGridView1.Rows(Baris).Cells(3).Value & "'where NoPinjam='" & DataGridView1.Rows(Baris).Cells(1).Value & "'"
					Cmd = New SqlCommand(Sqleditpjm, Conn)
					Cmd.ExecuteNonQuery()
				End If
			Next
			Call KondisiAwal()
			Call Otomatis()
			DataGridView1.Columns.Clear()
			DataGridView2.Columns.Clear()
			MsgBox("Transaksi Telah Berhasil Disimpan")
			Call TampilGrid1()
		End If
	End Sub
End Class