Imports System.Data.SqlClient
Public Class Form2
    Dim con As New SqlConnection
    Dim cmd As New SqlCommand
    Dim i As Integer
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.ConnectionString = "Data Source=LAPTOP-LP\SQLEXPRESS;Initial Catalog=corto;Integrated Security=True"
        If con.State = ConnectionState.Open Then
            con.Close()
        End If
        con.Open()
        disp_data()

        Label6.Text = Module1.Usuario
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim imc2 As String
        Dim peso, altura, imc As Double

        peso = TextBox1.Text
        altura = TextBox2.Text
        imc = peso / (altura * altura)
        Dim resultado As String = If(imc <= 18.5, "Bajo de peso",
                                        If(imc <= 24.9, "Peso normal",
                                        If(imc <= 29.9, "Sobrepeso",
                                        If(imc <= 34.9, "Obesidad grado I",
                                        If(imc <= 39.9, "Obesidad grado II", "Obesidad grado III (obesidad mórbida)")))))
        imc2 = imc.ToString("F4")
        TextBox3.Text = imc2
        Label5.Text = resultado
    End Sub
    Public Sub disp_data()
        cmd = con.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "SELECT * FROM IMC__202109710"
        cmd.ExecuteNonQuery()
        Dim dt As New DataTable()
        Dim da As New SqlDataAdapter(cmd)
        da.Fill(dt)
        DataGridView1.DataSource = dt
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        Try
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            con.Open()
            i = Convert.ToInt32(DataGridView1.SelectedCells.Item(0).Value.ToString())
            cmd = con.CreateCommand()
            cmd.CommandType = CommandType.Text
            cmd.CommandText = "select * from IMC__202109710 where IMC=" & i & ""
            cmd.ExecuteNonQuery()
            Dim dt As New DataTable()
            Dim da As New SqlDataAdapter(cmd)
            da.Fill(dt)
            Dim dr As SqlClient.SqlDataReader
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            While dr.Read
                TextBox3.Text = dr.GetString(1).ToString()
                Label5.Text = dr.GetString(2).ToString()
            End While
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        disp_data()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If con.State = ConnectionState.Open Then
            con.Close()
        End If
        con.Open()
        cmd = con.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "delete from IMC__202109710 where IMC ='" & TextBox3.Text & "'"
        cmd.ExecuteNonQuery()
        disp_data()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        cmd = con.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "insert into IMC__202109710 values('" + Label6.Text + "','" + TextBox3.Text + "','" & Label5.Text & "')"
        cmd.ExecuteNonQuery()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        Label5.Text = ""
        disp_data()
    End Sub
End Class