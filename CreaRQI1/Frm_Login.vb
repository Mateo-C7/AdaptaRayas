Public Class Frm_Login

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        acepta()
    End Sub

    Sub acepta()
        If TextBox1.Text = "" Then
            MsgBox("introduzca nombre de usuario")
        ElseIf TextBox2.Text = "" Then
            MsgBox("introduzca contraseña")
        Else
            CreaRQI.MyCommands.NomUsuario = TextBox1.Text
            CreaRQI.MyCommands.Passw = TextBox2.Text
            CreaRQI.MyCommands.sigue = True
            Me.Close()
        End If

    End Sub

    Sub verificatecla(ByVal e As Windows.Forms.KeyEventArgs)
        Try
            If e.KeyCode = Windows.Forms.Keys.Escape Then
                CreaRQI.MyCommands.sigue = False
                Me.Close()
            ElseIf e.KeyCode = Windows.Forms.Keys.Enter Then
                acepta()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Frm_Login_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        verificatecla(e)
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        verificatecla(e)
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        verificatecla(e)
    End Sub
End Class