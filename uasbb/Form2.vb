Imports System.Data.SqlClient
Public Class Form2
    Dim Conn As SqlConnection
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim LokasiDB As String
    'Sub Koneksi()
    '    LokasiDB = "<CONNECT TO YOUR DB>"
    '    Conn = New SqlConnection(LokasiDB)
    '    If Conn.State = ConnectionState.Closed Then Conn.Open()
    'End Sub

    Sub reload()
        Dim sa As String
        sa = "select * from dummyDum.dbo.sementara"
        Dim da As New SqlDataAdapter(sa, Conn)
        Dim dtMS As New DataTable
        da.Fill(dtMS)
        DataGridView1.DataSource = dtMS
    End Sub

    Sub reloadquerry()
        Dim sa, sb, sc As String
        sa = "select * from UASMeSc.dbo.ms_buku"
        sb = "select * from UASMeSc.dbo.ms_penjualan"
        sc = "select * from UASMeSc.dbo.ms_pembelian"
        Dim da As New SqlDataAdapter(sa, Conn)
        Dim db As New SqlDataAdapter(sb, Conn)
        Dim dc As New SqlDataAdapter(sc, Conn)
        Dim dtbk As New DataTable
        Dim dtpe As New DataTable
        Dim dtpj As New DataTable
        da.Fill(dtbk)
        DataGridView2.DataSource = dtbk
        db.Fill(dtpe)
        DataGridView3.DataSource = dtpe
        dc.Fill(dtpj)
        DataGridView4.DataSource = dtpj
    End Sub

    Sub reloadTb()
        ' set buku
        TextBox17.Text = DataGridView1.Rows(12).Cells(1).Value.ToString
        TextBox1.Text = DataGridView1.Rows(13).Cells(1).Value.ToString
        TextBox2.Text = DataGridView1.Rows(14).Cells(1).Value.ToString
        TextBox3.Text = DataGridView1.Rows(15).Cells(1).Value.ToString
        TextBox4.Text = DataGridView1.Rows(16).Cells(1).Value.ToString
        TextBox5.Text = DataGridView1.Rows(17).Cells(1).Value.ToString

        ' set penjualan
        TextBox6.Text = DataGridView1.Rows(0).Cells(1).Value.ToString
        TextBox7.Text = DataGridView1.Rows(1).Cells(1).Value.ToString
        TextBox8.Text = DataGridView1.Rows(2).Cells(1).Value.ToString
        TextBox9.Text = DataGridView1.Rows(3).Cells(1).Value.ToString
        TextBox10.Text = DataGridView1.Rows(4).Cells(1).Value.ToString

        'set pembelian
        TextBox11.Text = DataGridView1.Rows(5).Cells(1).Value.ToString
        TextBox12.Text = DataGridView1.Rows(6).Cells(1).Value.ToString
        TextBox13.Text = DataGridView1.Rows(7).Cells(1).Value.ToString
        TextBox14.Text = DataGridView1.Rows(8).Cells(1).Value.ToString
        TextBox15.Text = DataGridView1.Rows(9).Cells(1).Value.ToString
        TextBox16.Text = DataGridView1.Rows(10).Cells(1).Value.ToString
        TextBox18.Text = DataGridView1.Rows(11).Cells(1).Value.ToString


    End Sub

    Private Sub MatchingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MatchingToolStripMenuItem.Click
        Me.Close()
        Form1.Show()
    End Sub

    Private Sub MappingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MappingToolStripMenuItem.Click
        Me.Show()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Koneksi()
        reload()
        reloadTb()
        reloadquerry()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'MENGGANTI FROM MENJADI DIDAPAT DARI COL(0) DARI DATAGRIDVIEW
        Try
            Dim sa As New SqlCommand
            sa = New SqlCommand("
        insert into UASMeSc.dbo.ms_buku (kodebuku, namabuku, jenisbuku, namapengarang,namapenerbit, hargabuku)
        select a." & TextBox17.Text & ", c." & TextBox1.Text & " ,b." & TextBox2.Text & ", d." & TextBox3.Text & ", e." & TextBox4.Text & ", a." & TextBox5.Text & "
        From UASPenjualan.dbo.pe_jenis b, UASPenjualan.dbo.pe_buku a, UASPembelian.dbo.pj_buku c, UASPembelian.dbo.pj_pengarang d, UASPembelian.dbo.pj_penerbit e
        where c.nama_buku = a.judul and a.kode_jenis = b.kode_jenis and c.id_pengarang = d.id_pengarang and c.id_penerbit = e.id_penerbit ", Conn)
            sa.ExecuteNonQuery()

            Dim sb As New SqlCommand
            sb = New SqlCommand(
"insert into UASMeSc.dbo.ms_penjualan(nopenjualan,tanggalpenjualan,kodebuku,jumlahbeli,totalharga) 
select a." & TextBox6.Text & ", a." & TextBox7.Text & ", b." & TextBox8.Text & ", c." & TextBox9.Text & ", c." & TextBox10.Text & " 
from UASPenjualan.dbo.pe_penjualan a , UASPenjualan.dbo.pe_buku b , UASPenjualan.dbo.pe_dtlpenjualan c where a.nopenjualan = c.nopenjualan 
and c.kode_buku = b.kode_buku", Conn)
            sb.ExecuteNonQuery()

            Dim sc As New SqlCommand
            sc = New SqlCommand(
"insert into UASMeSc.dbo.ms_pembelian(nofaktur,kodebuku,jumlahpembelian,namasupplier,hargasatuan,totalhargapembelian, tanggalpembelian) 
select a." & TextBox11.Text & ", b." & TextBox12.Text & ", c." & TextBox13.Text & ", d." & TextBox14.Text & " , c." & TextBox15.Text & ", 
c." & TextBox16.Text & ", a." & TextBox18.Text & "
from UASPembelian.dbo.pj_pembelian a, UASPembelian.dbo.pj_dtlpembelian c, UASPembelian.dbo.pj_supplier d , UASPembelian.dbo.pj_buku e, UASPenjualan.dbo.pe_buku b 
where a.no_faktur = c.no_faktur and a.id_supplier = d.id_supplier  and e.id_buku = c.id_buku 
and SUBSTRING(b.kode_buku, 2, 7) = SUBSTRING(e.id_buku,2,7)", Conn)
            sc.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Something went wrong")
        End Try
        reloadquerry()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim sa, sb, sc As New SqlCommand
            sa = New SqlCommand("delete from UASMeSc.dbo.ms_buku", Conn)
            sa.ExecuteNonQuery()

            sb = New SqlCommand("delete from UASMeSc.dbo.ms_penjualan", Conn)
            sb.ExecuteNonQuery()

            sc = New SqlCommand("delete from UASMeSc.dbo.ms_pembelian", Conn)
            sc.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Something went wrong")
        End Try

        reloadquerry()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Dim sa As String
        sa = "insert into UASMeSc.dbo.ms_buku (kodebuku, namabuku, jenisbuku, namapengarang, namapenerbit,hargabuku) 
            select a." & TextBox17.Text & ", c." & TextBox1.Text & " , d." & TextBox2.Text & " , b." & TextBox3.Text & " , e." & TextBox4.Text & ", a." & TextBox5.Text & " 
            from UASPenjualan.dbo.pe_buku a ,UASPenjualan.dbo.pe_jenis b, UASPembelian.dbo.pj_buku c, UASPembelian.dbo.pj_pengarang d, UASPembelian.dbo.pj_penerbit e 
            where c." & TextBox1.Text & "  = a.judul and a.kode_jenis = b.kode_jenis and c.id_pengarang = d.id_pengarang and c.id_penerbit = e.id_penerbit"
        MsgBox(sa)
    End Sub

    Private Sub LaporanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LaporanToolStripMenuItem.Click
        Me.Close()
        Form3.Show()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs)

    End Sub
End Class