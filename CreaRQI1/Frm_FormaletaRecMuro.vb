Public Class Frm_FormaletaRecMuro
    Dim unidadactual As String
    Private Sub Frm_FormaletaRecMuro_Load(sender As Object, e As EventArgs) Handles Me.Load
        unidadactual = ComboBox1.Text
    End Sub
    Private Sub todosinvisibles()
        PictureBox2.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        PictureBox6.Visible = False
        PictureBox7.Visible = False
        PictureBox8.Visible = False
        PictureBox9.Visible = False
        PictureBox10.Visible = False
        PictureBox11.Visible = False
    End Sub
    Private Sub TextBox1_MouseHover(sender As Object, e As EventArgs) Handles TextBox1.MouseHover
        todosinvisibles()
        PictureBox2.Visible = True
    End Sub
    Private Sub TextBox2_MouseHover(sender As Object, e As EventArgs) Handles TextBox2.MouseHover
        todosinvisibles()
        PictureBox3.Visible = True
    End Sub
    Private Sub TextBox3_MouseHover(sender As Object, e As EventArgs) Handles TextBox3.MouseHover
        todosinvisibles()
        PictureBox4.Visible = True
    End Sub
    Private Sub TextBox4_MouseHover(sender As Object, e As EventArgs) Handles TextBox4.MouseHover
        todosinvisibles()
        PictureBox5.Visible = True
    End Sub
    Private Sub TextBox5_MouseHover(sender As Object, e As EventArgs) Handles TextBox5.MouseHover
        todosinvisibles()
        PictureBox6.Visible = True
    End Sub
    Private Sub TextBox6_MouseHover(sender As Object, e As EventArgs) Handles TextBox6.MouseHover
        todosinvisibles()
        PictureBox7.Visible = True
    End Sub
    Private Sub TextBox7_MouseHover(sender As Object, e As EventArgs) Handles TextBox7.MouseHover
        todosinvisibles()
        PictureBox8.Visible = True
    End Sub
    Private Sub TextBox8_MouseHover(sender As Object, e As EventArgs) Handles TextBox8.MouseHover
        todosinvisibles()
        PictureBox9.Visible = True
    End Sub
    Private Sub TextBox9_MouseHover(sender As Object, e As EventArgs) Handles TextBox9.MouseHover
        todosinvisibles()
        PictureBox10.Visible = True
    End Sub
    Private Sub TextBox10_MouseHover(sender As Object, e As EventArgs) Handles TextBox10.MouseHover
        todosinvisibles()
        PictureBox11.Visible = True
    End Sub



    Private Sub Label1_MouseHover(sender As Object, e As EventArgs) Handles Label1.MouseHover
        todosinvisibles()
        PictureBox2.Visible = True
    End Sub
    Private Sub Label2_MouseHover(sender As Object, e As EventArgs) Handles Label2.MouseHover
        todosinvisibles()
        PictureBox3.Visible = True
    End Sub
    Private Sub Label3_MouseHover(sender As Object, e As EventArgs) Handles Label3.MouseHover
        todosinvisibles()
        PictureBox4.Visible = True
    End Sub
    Private Sub Label4_MouseHover(sender As Object, e As EventArgs) Handles Label4.MouseHover
        todosinvisibles()
        PictureBox5.Visible = True
    End Sub
    Private Sub Label5_MouseHover(sender As Object, e As EventArgs) Handles Label5.MouseHover
        todosinvisibles()
        PictureBox6.Visible = True
    End Sub
    Private Sub Label6_MouseHover(sender As Object, e As EventArgs) Handles Label6.MouseHover
        todosinvisibles()
        PictureBox7.Visible = True
    End Sub
    Private Sub Label7_MouseHover(sender As Object, e As EventArgs) Handles Label7.MouseHover
        todosinvisibles()
        PictureBox8.Visible = True
    End Sub
    Private Sub Label8_MouseHover(sender As Object, e As EventArgs) Handles Label8.MouseHover
        todosinvisibles()
        PictureBox9.Visible = True
    End Sub
    Private Sub Label9_MouseHover(sender As Object, e As EventArgs) Handles Label9.MouseHover
        todosinvisibles()
        PictureBox10.Visible = True
    End Sub
    Private Sub Label10_MouseHover(sender As Object, e As EventArgs) Handles Label10.MouseHover
        todosinvisibles()
        PictureBox11.Visible = True
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox1.BackColor = Drawing.Color.White
        TextBox2.BackColor = Drawing.Color.White
        TextBox3.BackColor = Drawing.Color.White
        TextBox4.BackColor = Drawing.Color.White
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.BackColor = Drawing.Color.White
        TextBox2.BackColor = Drawing.Color.White
        TextBox3.BackColor = Drawing.Color.White
        TextBox4.BackColor = Drawing.Color.White
        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox4.Text = "" Then
            If TextBox1.Text = "" Then
                TextBox1.BackColor = Drawing.Color.Yellow
            End If
            If TextBox2.Text = "" Then
                TextBox2.BackColor = Drawing.Color.Yellow
            End If
            If TextBox4.Text = "" Then
                TextBox4.BackColor = Drawing.Color.Yellow
            End If
            MsgBox("se requiere informacion en los campos resaltados",, "falta informacion")
        Else
            TextBox3.Text = Math.Round((360 * CDbl(TextBox2.Text)) / (2 * Math.PI * CDbl(TextBox1.Text)), 1)
            TextBox5.Text = CDbl(TextBox3.Text) - CDbl(TextBox4.Text)
            TextBox6.Text = Math.Floor((CDbl(TextBox5.Text) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox7.Text = Math.Floor(((CDbl(TextBox5.Text) - 5.4) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox8.Text = Math.Floor(((CDbl(TextBox5.Text) - 6.5) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox9.Text = Math.Floor(((CDbl(TextBox5.Text) - 8) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox10.Text = Math.Floor(((CDbl(TextBox5.Text) - 10) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox11.Text = "Modular piezas externas rectas con platinas verticales de 90° a un ancho maximo de " & TextBox2.Text +
            unidadactual & ", espesor de muro de " & TextBox4.Text & unidadactual & "." & vbCrLf & "Si las piezas internas son de 5.4cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox7.Text), 1)) +
            unidadactual & " en el ancho." & vbCrLf & "si las piezas internas son de 6.5cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox8.Text), 1)) +
            unidadactual & " en el ancho." & vbCrLf & "si las piezas internas son de 8cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox9.Text), 1)) +
            unidadactual & " en el ancho." & vbCrLf & "si las piezas internas son de 10cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox10.Text), 1)) +
            unidadactual & " en el ancho."
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.BackColor = Drawing.Color.White
        TextBox2.BackColor = Drawing.Color.White
        TextBox3.BackColor = Drawing.Color.White
        TextBox4.BackColor = Drawing.Color.White
        If TextBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Then
            If TextBox1.Text = "" Then
                TextBox1.BackColor = Drawing.Color.Yellow
            End If
            If TextBox3.Text = "" Then
                TextBox3.BackColor = Drawing.Color.Yellow
            End If
            If TextBox4.Text = "" Then
                TextBox4.BackColor = Drawing.Color.Yellow
            End If
            MsgBox("se requiere informacion en los campos resaltados",, "falta informacion")
        Else
            TextBox2.Text = Math.Round(CDbl(TextBox3.Text) * 2 * Math.PI / 360 * CDbl(TextBox1.Text), 1)
            TextBox5.Text = CDbl(TextBox3.Text) - CDbl(TextBox4.Text)
            TextBox6.Text = Math.Floor((CDbl(TextBox5.Text) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox7.Text = Math.Floor(((CDbl(TextBox5.Text) - 5.4) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox8.Text = Math.Floor(((CDbl(TextBox5.Text) - 6.5) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox9.Text = Math.Floor(((CDbl(TextBox5.Text) - 8) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox10.Text = Math.Floor(((CDbl(TextBox5.Text) - 10) * Math.PI * 2 / 360 * CDbl(TextBox1.Text)) * 10) / 10
            TextBox11.Text = "Modular piezas externas rectas con platinas verticales de 90° a un ancho maximo de " & TextBox2.Text +
            unidadactual & ", espesor de muro de " & TextBox4.Text & unidadactual & "." & vbCrLf & "Si las piezas internas son de 5.4cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox7.Text), 1)) +
            unidadactual & " en el ancho." & vbCrLf & "si las piezas internas son de 6.5cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox8.Text), 1)) +
            unidadactual & " en el ancho." & vbCrLf & "si las piezas internas son de 8cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox9.Text), 1)) +
            unidadactual & " en el ancho." & vbCrLf & "si las piezas internas son de 10cm se deben modular restando " & CStr(Math.Round(CDbl(TextBox2.Text) - CDbl(TextBox10.Text), 1)) +
            unidadactual & " en el ancho."
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = unidadactual Then
            'si seleccionó la misma que tenía, no haga calculos
        Else
            TextBox11.Text = ""
            If ComboBox1.Text = "mm" Then
                unidadactual = ComboBox1.Text
                cambiarunidades("mm")
                Dim cajas() As Windows.Forms.TextBox = {TextBox2, TextBox3, TextBox4, TextBox5, TextBox6, TextBox7, TextBox8, TextBox9, TextBox10}
                For i = 0 To cajas.Length - 1
                    multiplicapordiez(cajas(i))
                Next
            ElseIf ComboBox1.Text = "cm" Then
                unidadactual = ComboBox1.Text
                cambiarunidades("cm")
                Dim cajas() As Windows.Forms.TextBox = {TextBox2, TextBox3, TextBox4, TextBox5, TextBox6, TextBox7, TextBox8, TextBox9, TextBox10}
                For i = 0 To cajas.Length - 1
                    dividepordiez(cajas(i))
                Next
            End If
        End If

    End Sub
    Private Sub dividepordiez(ByRef textbox As Windows.Forms.TextBox)
        If textbox.Text = "" Then
        Else
            textbox.Text = CDbl(textbox.Text) / 10
        End If
    End Sub
    Private Sub multiplicapordiez(ByRef textbox As Windows.Forms.TextBox)
        If textbox.Text = "" Then
        Else
            textbox.Text = CDbl(textbox.Text) * 10
        End If
    End Sub
    Private Sub cambiarunidades(ByVal unidad As String)
        Label12.Text = unidad
        Label13.Text = unidad
        Label14.Text = unidad
        Label15.Text = unidad
        Label16.Text = unidad
        Label17.Text = unidad
        Label18.Text = unidad
        Label19.Text = unidad
        Label20.Text = unidad
    End Sub
End Class