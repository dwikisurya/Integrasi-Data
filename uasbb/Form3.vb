Imports System.Data.SqlClient

Public Class Form3
    Dim Conn As SqlConnection
    Dim LokasiDB As String

    'Sub Koneksi()
    '    LokasiDB = "<CONNECT TO YOUR DB>"
    '    Conn = New SqlConnection(LokasiDB)
    '    If Conn.State = ConnectionState.Closed Then Conn.Open()
    'End Sub
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Koneksi()
        'loadpenjualan()
        'loadlaporankeuntungan()
    End Sub


    Sub loadpenjualan()
        'Dim sa As String
        'sa = "Select a.kodebuku, b.namabuku, a.tanggalpenjualan, a.jumlahbeli, a.totalharga
        '      from UASMeSc.dbo.ms_penjualan a , UASMeSc.dbo.ms_buku b where a.kodebuku = b.kodebuku"
        'Dim da As New SqlDataAdapter(sa, Conn)
        'Dim dtLaporanPenjualan As New DataTable
        'da.Fill(dtLaporanPenjualan)
        'DataGridView1.DataSource = dtLaporanPenjualan
    End Sub

    'Sub loadpembelian()
    '    Dim sb As String
    '    sb = "Select a.kodebuku, b.namabuku, a.tanggalpembelian , a.jumlahpembelian,a.namasupplier ,a.totalhargapembelian
    '          from UASMeSc.dbo.ms_pembelian a , UASMeSc.dbo.ms_buku b where a.kodebuku = b.kodebuku"
    '    Dim db As New SqlDataAdapter(sb, Conn)
    '    Dim dtLaporanPembelian As New DataTable
    '    db.Fill(dtLaporanPembelian)
    '    DataGridView2.DataSource = dtLaporanPembelian
    'End Sub

    Sub loadlaporankeuntungan()
        Try
            Dim sc As String
            sc = "select distinct b.kodebuku, a.namabuku, b.tanggalpenjualan, null as 'tanggalpembelian', b.totalharga, '0' as 'totalhargapembelian'
from ms_buku a 
inner join ms_penjualan b on b.kodebuku = a.kodebuku
inner join ms_pembelian c on c.kodebuku = a.kodebuku
where month(b.tanggalpenjualan) = " & ComboBox1.Text & "

            union all

select distinct c.kodebuku, a.namabuku, null ,c.tanggalpembelian, '0' as 'totalharga ', c.totalhargapembelian
from ms_buku a
inner join ms_pembelian c on c.kodebuku = a.kodebuku
inner join ms_penjualan b on c.kodebuku = b.kodebuku
where month(c.tanggalpembelian) = " & ComboBox1.Text & ""
            Dim dc As New SqlDataAdapter(sc, Conn)
            Dim dtLaporanKeuntungan As New DataTable
            dc.Fill(dtLaporanKeuntungan)
            DataGridView3.DataSource = dtLaporanKeuntungan
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub



    Private Sub MatchingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MatchingToolStripMenuItem.Click
        Me.Close()
        Form1.Show()
    End Sub

    Private Sub MappingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MappingToolStripMenuItem.Click
        Me.Close()
        Form2.Show()
    End Sub

    '    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '        Try
    '            Dim sb As String
    '            sb = "Select a.kodebuku, b.namabuku, a.tanggalpembelian , a.jumlahpembelian,a.namasupplier ,a.totalhargapembelian
    'from UASMeSc.dbo.ms_pembelian a , UASMeSc.dbo.ms_buku b where a.kodebuku = b.kodebuku and MONTH(a.tanggalpembelian) = " & ComboBox1.Text & ""
    '            Dim db As New SqlDataAdapter(sb, Conn)
    '            Dim dtLaporanPembelian As New DataTable
    '            db.Fill(dtLaporanPembelian)
    '            DataGridView2.DataSource = dtLaporanPembelian
    '        Catch ex As Exception
    '            MsgBox("error")
    '        End Try

    '        Try
    '            Dim sa As String
    '            sa = "Select a.kodebuku, b.namabuku, a.tanggalpenjualan, a.jumlahbeli, a.totalharga
    '              from UASMeSc.dbo.ms_penjualan a , UASMeSc.dbo.ms_buku b where a.kodebuku = b.kodebuku and MONTH(a.tanggalpenjualan) = " & ComboBox1.Text & ""
    '            Dim da As New SqlDataAdapter(sa, Conn)
    '            Dim dtLaporanPenjualan As New DataTable
    '            da.Fill(dtLaporanPenjualan)
    '            DataGridView1.DataSource = dtLaporanPenjualan
    '        Catch ex As Exception
    '            MsgBox("error")
    '        End Try
    '        loadlaporankeuntungan()
    '    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Dim sb As String
            sb = "Select a.kodebuku, b.namabuku, a.tanggalpembelian , a.jumlahpembelian,a.namasupplier ,a.totalhargapembelian
from UASMeSc.dbo.ms_pembelian a , UASMeSc.dbo.ms_buku b where a.kodebuku = b.kodebuku and MONTH(a.tanggalpembelian) = " & ComboBox1.Text & ""
            Dim db As New SqlDataAdapter(sb, Conn)
            Dim dtLaporanPembelian As New DataTable
            db.Fill(dtLaporanPembelian)
            DataGridView2.DataSource = dtLaporanPembelian
        Catch ex As Exception
            MsgBox("error")
        End Try

        Try
            Dim sa As String
            sa = "Select a.kodebuku, b.namabuku, a.tanggalpenjualan, a.jumlahbeli, a.totalharga
              from UASMeSc.dbo.ms_penjualan a , UASMeSc.dbo.ms_buku b where a.kodebuku = b.kodebuku and MONTH(a.tanggalpenjualan) = " & ComboBox1.Text & ""
            Dim da As New SqlDataAdapter(sa, Conn)
            Dim dtLaporanPenjualan As New DataTable
            da.Fill(dtLaporanPenjualan)
            DataGridView1.DataSource = dtLaporanPenjualan
        Catch ex As Exception
            MsgBox("error")
        End Try
        loadlaporankeuntungan()
    End Sub
End Class