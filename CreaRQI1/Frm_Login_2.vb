Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports CreaRQI.CapaDatos

Public Class Frm_Login_2
    Dim Conn As ConexionBD = New ConexionBD()
    Public darusuario As String
    Public darpass As String
    Public nombre As String
    Public validez As Boolean
    Public iduser As String
    Public Str_Con_Bd As String = ConexionBD.getStringConexion()
    Private Sub Frm_Login_2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'carganombres()
        If TextBox1.Text <> "" And TextBox2.Text <> "" Then
            verifica2(TextBox2.Text, TextBox1.Text)
        Else
            ComboBox1.Visible = False
        End If

    End Sub
    Public Sub carganombres()
        Dim Cons As String = "SELECT usuario.usu_passwd, usuario.usu_activo, area.are_id, empleado.emp_nombre_mayusculas, empleado.emp_apellidos_mayusculas " +
                "FROM  usuario INNER JOIN " +
                "empleado ON usuario.usu_emp_usu_num_id = empleado.emp_usu_num_id INNER JOIN " +
                "area ON empleado.emp_area_id = area.are_id " +
                "WHERE (usuario.usu_activo = 1) AND (area.are_id = 17) AND (NOT (empleado.emp_nombre = 'Sin nombre')) " +
                "ORDER BY empleado.emp_nombre_mayusculas"
        Using connection As New SqlConnection(Str_Con_Bd)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                ComboBox1.Items.Clear()
                Do While reader.Read()
                    ComboBox1.Items.Add(reader.GetValue(3) & " " & reader.GetValue(4))
                Loop
            End If
            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        'If Asc(e.KeyChar) = 13 Then
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then
            'verifica(ComboBox1.Text, TextBox1.Text)
            verifica2(TextBox2.Text, TextBox1.Text)
            darusuario = TextBox2.Text
            darpass = TextBox1.Text
            Me.Hide()
        End If
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        'If Asc(e.KeyChar) = 13 Then
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then
            TextBox1.Focus()
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'verifica(ComboBox1.Text, TextBox1.Text)
        verifica2(TextBox2.Text, TextBox1.Text)
        darusuario = TextBox2.Text
        darpass = TextBox1.Text
        Me.Hide()
    End Sub
    Public Sub verifica2(ByVal Usuario As String, pass As String)
        Dim Cons As String = "SELECT usuario.usu_login, usuario.usu_passwd, empleado.emp_nombre_mayusculas, empleado.emp_apellidos_mayusculas, usuario.usu_emp_usu_num_id " +
                                "FROM  usuario INNER JOIN " +
                                "empleado ON usuario.usu_emp_usu_num_id = empleado.emp_usu_num_id " +
                                "WHERE (NOT (empleado.emp_nombre_mayusculas = 'Sin nombre')) " +
                                "ORDER BY usuario.usu_login"
        Using connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(Cons, connection)
            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            validez = False
            While reader.Read()
                Dim a = reader("usu_login").ToString
                If reader("usu_login").ToString = Usuario Then
                    If reader("usu_passwd").ToString = pass Then
                        validez = True
                        nombre = reader("emp_nombre_mayusculas").ToString & " " & reader("emp_apellidos_mayusculas").ToString
                        iduser = reader("usu_emp_usu_num_id").ToString
                        reader.Close()
                        reader = Nothing
                        Me.Hide()
                        Exit While
                    End If
                End If
            End While
            If validez = True Then

            Else
                MsgBox("usuario o clave no es correcta")
                TextBox2.Focus()
            End If
        End Using
    End Sub
    Public Sub verifica(ByRef IndexUsuario As String, pass As String)
        validez = False
        Dim Cons As String = "SELECT usuario.usu_passwd, usuario.usu_activo, area.are_id, empleado.emp_nombre_mayusculas, empleado.emp_apellidos_mayusculas " +
                "FROM  usuario INNER JOIN " +
                "empleado ON usuario.usu_emp_usu_num_id = empleado.emp_usu_num_id INNER JOIN " +
                "area ON empleado.emp_area_id = area.are_id " +
                "WHERE (usuario.usu_activo = 1) AND (area.are_id = 17) AND (NOT (empleado.emp_nombre = 'Sin nombre')) AND (RTRIM(empleado.emp_nombre_mayusculas) + ' ' + RTRIM(empleado.emp_apellidos_mayusculas) = '" & IndexUsuario & "') " +
                "ORDER BY empleado.emp_nombre_mayusculas "
        '"OFFSET  " & IndexUsuario & " ROWS "
        '"FETCH NEXT 1 ROWS ONLY"
        Using connection As New SqlConnection(Str_Con_Bd)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            reader.Read()
            'Dim pasw As String = reader.GetValue(0)
            If pass = reader.GetValue(0) Then
                validez = True
                nombre = reader.GetValue(3) & " " & reader.GetValue(4)
                command = Nothing
                reader.Close()
                reader = Nothing
                Me.Close()
            Else
                MsgBox("La clave no es correcta")
            End If
        End Using

    End Sub

End Class