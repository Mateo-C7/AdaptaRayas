Imports System.Data.SqlClient
Imports CreaRQI.CapaDatos

Public Class Frm_GrupoPlanta

    Dim Conn As ConexionBD = New ConexionBD 'Instanciamos la clase de conexion
    Public publicitemgrupoplantaid As String
    Public nom As String
    Public Str_Con_Bd As String = ConexionBD.getStringConexion()
    Private Sub Frm_GrupoPlanta_Load(sender As Object, e As EventArgs) Handles Me.Load
        publicitemgrupoplantaid = ""
        Me.Text = "Seleccione Grupo Planta para: " & nom
        Dim cons As String
        cons = "SELECT descripcion FROM item_grupo WHERE (activo = 1) ORDER BY descripcion"
        Using Connection As New SqlConnection(Str_Con_Bd)
            Dim command As SqlCommand = New SqlCommand(cons, Connection)
            Connection.Open()
            Dim reader As SqlDataReader
            reader = command.ExecuteReader()

            Do While reader.Read()
                ComboBox1.Items.Add(reader.GetValue(0))
            Loop
            'ComboBox1.Items.Add("NUEVO")
        End Using
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        itemgrupoplantaid(ComboBox1.Text)
        Me.Hide()
    End Sub
    Sub itemgrupoplantaid(ByVal descgrupoplanta As String)
        Dim itemgrupoid As Integer
        Dim cons As String
        Try

            cons = "SELECT item_grupo_id FROM item_grupo WHERE (descripcion = '" & descgrupoplanta & "')"
            Using Connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, Connection)
                Connection.Open()
                Dim reader As SqlDataReader
                reader = command.ExecuteReader()
                If reader.HasRows Then
                    reader.Read()
                    itemgrupoid = reader.GetValue(0)
                End If
            End Using
            cons = "SELECT item_grupo_planta_id FROM item_grupo_planta WHERE (item_grupo_id = " & itemgrupoid & ") AND (planta_id = 3)"
            Using Connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, Connection)
                Connection.Open()
                Dim reader As SqlDataReader
                reader = command.ExecuteReader()
                If reader.HasRows Then
                    reader.Read()
                    publicitemgrupoplantaid = reader.GetValue(0)
                End If
            End Using
            cons = "INSERT INTO Accesorios_Cod_GrupoPlanta (Nomenclatura,IdItemGrupoPlanta,IdPlanta) VALUES ('" & nom & "'," & publicitemgrupoplantaid & ",3)"
            Using connection As New SqlConnection(Str_Con_Bd)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                command.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            MsgBox("No hay GrupoPlanta para " & nom)
        End Try

    End Sub
End Class