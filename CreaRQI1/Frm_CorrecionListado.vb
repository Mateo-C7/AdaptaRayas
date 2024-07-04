Public Class Frm_CorrecionListado
    Public mensajes As String()
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For i = 0 To UBound(mensajes)
            TextBox1.AppendText(mensajes(i) & vbCrLf)
        Next
    End Sub
End Class