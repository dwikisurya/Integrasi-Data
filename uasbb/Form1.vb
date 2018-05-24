Imports System.Data.SqlClient

Public Class Form1
    Dim Conn As SqlConnection
    Dim da As SqlDataAdapter
    Dim dt As DataTable
    Dim LokasiDB As String

    'Sub Koneksi()
    '    LokasiDB = "<CONNECT TO YOUR DB>"
    '    Conn = New SqlConnection(LokasiDB)
    '    If Conn.State = ConnectionState.Closed Then Conn.Open()
    'End Sub

    Sub comboBoxTabel()
        Dim dtMediatedSchema = Conn.GetSchema("TABLES")
        ComboBox1.DataSource = dtMediatedSchema
        ComboBox1.DisplayMember = "table_name"
        ComboBox1.ValueMember = "table_name"
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Koneksi()
        comboBoxTabel()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        Dim da As New SqlDataAdapter("select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '" & ComboBox1.SelectedValue.ToString() & "'", Conn)
        Dim dtKolomMediatedSchema As New DataTable
        da.Fill(dtKolomMediatedSchema)
        ComboBox2.DisplayMember = "COLUMN_NAME"
        ComboBox2.DataSource = dtKolomMediatedSchema
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Threshold tidak boleh kosong")
        End If
        If RadioButton1.Checked Then
            Dim sa As String
            Dim sb As String
            TextBox2.Text = ""
            TextBox3.Text = 0
            TextBox4.Text = ""
            TextBox5.Text = 0
            sa = "Select d.TABLE_NAME as NamaTabel, e.COLUMN_NAME as NamaKolom, round(dbo.JaroWinklerStringSimilarity('" & ComboBox2.Text & "', e.COLUMN_NAME),1) as Similarity from UASPenjualan.INFORMATION_SCHEMA.COLUMNS e 
                    left outer join UASPenjualan.INFORMATION_SCHEMA.TABLES d on d.TABLE_CATALOG = e.TABLE_CATALOG and d.TABLE_SCHEMA = e.TABLE_SCHEMA and d.TABLE_NAME = e.TABLE_NAME where e.TABLE_NAME not in ('sysdiagrams') and round(dbo.JaroWinklerStringSimilarity('" & ComboBox2.Text & "', e.COLUMN_NAME),1) >= " & Convert.ToDecimal(TextBox1.Text) & "order by 3 desc"

            sb = "Select d.TABLE_NAME as NamaTabel, e.COLUMN_NAME as NamaKolom, round(dbo.JaroWinklerStringSimilarity('" & ComboBox2.Text & "', e.COLUMN_NAME),1) as Similarity from UASPembelian.INFORMATION_SCHEMA.COLUMNS e 
                    left outer join UASPembelian.INFORMATION_SCHEMA.TABLES d on d.TABLE_CATALOG = e.TABLE_CATALOG and d.TABLE_SCHEMA = e.TABLE_SCHEMA and d.TABLE_NAME = e.TABLE_NAME where e.TABLE_NAME not in ('sysdiagrams') and round(dbo.JaroWinklerStringSimilarity('" & ComboBox2.Text & "', e.COLUMN_NAME),1) >= " & Convert.ToDecimal(TextBox1.Text) & "order by 3 desc"

            Dim da As New SqlDataAdapter(sa, Conn)
            Dim db As New SqlDataAdapter(sb, Conn)
            Dim dtTabelPenjualan As New DataTable
            Dim dttabelpembelian As New DataTable
            da.Fill(dtTabelPenjualan)
            db.Fill(dttabelpembelian)
            DataGridView1.DataSource = dtTabelPenjualan
            DataGridView2.DataSource = dttabelpembelian

            If DataGridView1.Rows.Count > 1 And DataGridView1.Rows(0).Cells(2).Value <> 0 And DataGridView1.Rows(0).Cells(2).Value >= Convert.ToDecimal(TextBox1.Text) Then
                TextBox6.Text = DataGridView1.Rows(0).Cells(0).Value.ToString
                TextBox2.Text = DataGridView1.Rows(0).Cells(1).Value.ToString
                TextBox3.Text = DataGridView1.Rows(0).Cells(2).Value.ToString
            End If

            If DataGridView2.Rows.Count > 1 And DataGridView2.Rows(0).Cells(2).Value <> 0 And DataGridView2.Rows(0).Cells(2).Value >= Convert.ToDecimal(TextBox1.Text) Then
                TextBox7.Text = DataGridView2.Rows(0).Cells(0).Value.ToString
                TextBox4.Text = DataGridView2.Rows(0).Cells(1).Value.ToString
                TextBox5.Text = DataGridView2.Rows(0).Cells(2).Value.ToString
            End If

        End If
        'If RadioButton2.Checked Then
        '    Dim sa As String
        '    sa = "select *, round(dbo.JaroWinklerStringSimilarity('" & ComboBox2.Text & "', d.COLUMN_NAME),1) as Similarity from " & ComboBox3.Text & ".INFORMATION_SCHEMA.COLUMNS d
        'where d.COLUMN_NAME NOT IN ('definition', 'diagram_id', 'name', 'version', 'principal_id' )and round(dbo.JaroWinklerStringSimilarity('" & ComboBox2.Text & "', d.COLUMN_NAME),1) > " & Convert.ToDecimal(TextBox1.Text) & ""
        '    Dim da As New SqlDataAdapter(sa, Conn)
        '    Dim dtTabel As New DataTable
        '    da.Fill(dtTabel)
        '    DataGridView1.DataSource = dtTabel
        'End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Convert.ToDecimal(TextBox3.Text) > Convert.ToDecimal(TextBox5.Text) Then
            Dim sq As New SqlCommand
            sq = New SqlCommand("insert into dummyDum.dbo.sementara(namatabel,namakolom, similarity)  values('" & TextBox6.Text & "','" & TextBox2.Text & "', '" & TextBox3.Text & "')", Conn)
            sq.ExecuteNonQuery()
        End If
        If Convert.ToDecimal(TextBox3.Text) < Convert.ToDecimal(TextBox5.Text) Then
            Dim sq As New SqlCommand
            sq = New SqlCommand("insert into dummyDum.dbo.sementara(namatabel,namakolom, similarity)  values('" & TextBox7.Text & "','" & TextBox4.Text & "', '" & TextBox5.Text & "')", Conn)
            sq.ExecuteNonQuery()
        End If
    End Sub

    Private Sub MatchingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MatchingToolStripMenuItem.Click
        Me.Show()
    End Sub

    Private Sub MappingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MappingToolStripMenuItem.Click
        Form2.Show()
    End Sub

    Private Sub LaporanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LaporanToolStripMenuItem.Click
        Form3.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim sa As New SqlCommand
            sa = New SqlCommand("delete from dummyDum.dbo.sementara", Conn)
            sa.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Failll")
        End Try
    End Sub
End Class
