Public Class Frm_SeleccionarPlantaACC
    Public planta As String
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        planta = ComboBox1.Text
        Me.Hide()
    End Sub
End Class