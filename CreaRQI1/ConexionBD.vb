Imports System
Imports System.Configuration
Imports System.Data
Imports System.Data.SqlClient

Namespace CapaDatos 'VB no genera el namespace por defecto, asi que hay que agregarlo manual
    Public Class ConexionBD

        Private cnnSqlS As SqlConnection
        Private comando As SqlCommand

        Public Function conectar() As Integer

            cnnSqlS = New SqlConnection()
            cnnSqlS.ConnectionString = getStringConexion()

            Try

                cnnSqlS.Open()
                Return 1

            Catch ex As Exception

                RegistraExcepcion("SYS", "BdDatos.Conectar", ex.Message.Replace("'", ""))
                Return 0

            End Try

        End Function

        Public Function desconectar() As Integer

            Try

                cnnSqlS.Close()
                Return 1

            Catch ex As Exception

                Return 0

            End Try

        End Function

        Public Function ejecutarSql(ByVal sql As String) As Integer

            Dim filas_afectadas As Integer = 0
            Dim MySql As String

            If conectar() > 0 Then
                MySql = " BEGIN TRANSACTION " & sql & " COMMIT;"
                comando = New SqlCommand(MySql, cnnSqlS)
                comando.CommandTimeout = 600
                filas_afectadas = comando.ExecuteNonQuery()
            End If

            desconectar()
            Return (filas_afectadas)

        End Function

        Public Shared Function getStringConexion() As String

            'No hay que agregarle a esta var el .ToString() aparentemente por el tipo de dato (string) el lo hace implicitamente
            Dim TpCon As String
            Dim CadenaCon As String
            Dim IpSource As String
            Dim Catalog As String

            TpCon = [CONST].TC

            Select Case TpCon
                Case "R"
                    Catalog = "Forsa"
                    IpSource = "10.75.131.2"
                    CadenaCon = "data source=" & IpSource & "; persist security info=False;initial catalog=" & Catalog & "; user id=forsa; password=forsa2006"
                Case "P"
                    Catalog = "ForsaMetro"
                    IpSource = "172.21.0.39"
                    CadenaCon = "data source=" & IpSource & "; persist security info=False;initial catalog=" & Catalog & "; user id=forsa; password=forsa2006"
                Case Else
                    Catalog = "ForsaMetro"
                    IpSource = "172.21.0.39"
                    CadenaCon = "data source=" & IpSource & "; persist security info=False;initial catalog=" & Catalog & "; user id=forsa; password=forsa2006"
            End Select

            Return CadenaCon

        End Function

        Public Function SelReg(ByVal sql As String) As DataRow
            Dim reg As DataRow = Nothing

            Try
                Dim dtRow As DataTable = CargarTabla(sql)

                If dtRow IsNot Nothing Then

                    For Each rg As DataRow In dtRow.Rows
                        reg = rg
                    Next
                Else
                    RegistraExcepcion("SYS", "BdDatos.SelReg", "Ocorrio un error en la ejecuación de la consulta", sql)
                End If

                Return reg
            Catch e As Exception
                RegistraExcepcion("SYS", "BdDatos.SelReg", e.Message.Replace("'", ""), sql)
                Return Nothing
            End Try
        End Function

        Public Function CargarTabla(ByVal sql As String) As DataTable
            Dim adapter As SqlDataAdapter
            Dim comando As SqlCommand
            Dim tabla As DataTable
            tabla = New DataTable()

            Try

                If conectar() > 0 Then

                    If cnnSqlS IsNot Nothing Then
                        adapter = New SqlDataAdapter()
                        If cnnSqlS.State = ConnectionState.Closed Then cnnSqlS.Open()
                        comando = New SqlCommand(sql, cnnSqlS)
                        comando.CommandTimeout = 600
                        adapter.SelectCommand = comando
                        adapter.FillSchema(tabla, SchemaType.Source)
                        adapter.Fill(tabla)
                        desconectar()
                    End If

                    Return tabla
                Else
                    RegistraExcepcion("SYS", "BdDatos.CargarTabla", "No fue posible abrir la conexión", sql)
                    Return Nothing
                End If

            Catch e As Exception
                RegistraExcepcion("SYS", "BdDatos.CargarTabla", e.Message.Replace("'", ""), sql)
                Return Nothing
            End Try
        End Function

        Public Function RegistraExcepcion(ByVal user As String, ByVal stModulo As String, ByVal stMes As String, ByVal Optional stSql As String = "") As Integer
            Dim sqlSel As String
            Dim sqlErr As String
            sqlErr = stSql.Replace("'", " ")

            If String.IsNullOrEmpty(user) Then
                user = "CARGA_SIIF"
            End If

            sqlSel = " INSERT INTO LOGSIIF (Usuario,Fecha,Origen,Detalle,Sistema,Query,SistVers) " & " VALUES ( '" & user & "' , SYSDATETIME(),'" & stModulo & "','" & stMes & "', 'CARGA_SIIF','" & sqlErr & "','" & [CONST].VER & "')"
            Return ejecutarSql(sqlSel)
        End Function
    End Class
End Namespace
