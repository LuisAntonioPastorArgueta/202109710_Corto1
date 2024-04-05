Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim usuario, contra As String
        usuario = TextBox1.Text
        contra = TextBox2.Text

        If usuario = "LUIS P" And contra = "123" Then
            Module1.Usuario = TextBox1.Text
            Form2.Show()
            Me.Hide()
        Else
            MsgBox("Datos incorrectos, prueba de nuevo")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub
End Class

