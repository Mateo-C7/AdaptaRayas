Imports System.Configuration
Imports CreaRQI.CapaDatos
Imports CreaRQI.CapaDatos.ConexionBD
Public Class Frm_Prueba

    Dim Conn As ConexionBD = New ConexionBD

    Private Sub Btn_Consultar_Click(sender As Object, e As EventArgs) Handles Btn_Consultar.Click

        Dim regVal As DataRow

        Dim Sql As String = "SELECT Nombres, Apellidos FROM VW_SM_EMPLEADOS WHERE Id = " + Txt_Entrada.Text

        regVal = Conn.SelReg(Sql)

        If regVal IsNot Nothing Then

            Label1.Text = regVal("Nombres") + " " + regVal("Apellidos")

        End If

        'Label1.Text = Txt_Entrada.Text.Trim()



    End Sub

    Private Sub Frm_Prueba_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label1.Text = ""

    End Sub
End Class