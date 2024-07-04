Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms
Imports CreaRQI.CapaDatos



Public Class Frm_ListFormaletas

    Dim Conn As ConexionBD = New ConexionBD 'Instanciamos la clase de conexion
    Dim Ord As Orden = New Orden 'Instanciamos la clase Orden
    Dim Util As Utility = New Utility 'Instanciamos la clase Utility
    Dim combinar As Boolean = False
    Dim primercombi As Boolean = True
    Dim listacombi(,) As String
    'Dim cadenaconexion = "data source = 172.21.224.130 ; persist security info=False;initial catalog = forsa ; user id=forsa; password=forsa2006 "
    Public cadenaconexion As String = ConexionBD.getStringConexion()
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboTipoOrden.SelectedIndexChanged
        Carga_Combo2(cboTipoOrden.SelectedItem)
    End Sub
    Public Sub Carga_Combo2(ByRef T_SolX As String) ', Orig_ As String)

        Dim Cons As String = "Select Orden.ano, Orden.Numero, Orden.letra, Orden.Abierta, Orden.Tipo_Of, Orden.Id_Of_P FROM Orden With (nolock) GROUP BY " +
            "Orden.ano, Orden.Numero, Orden.letra, Orden.Abierta, Orden.Tipo_Of, Orden.Id_Of_P HAVING (((Orden.Tipo_Of) " +
            "= '" & T_SolX & "')"
        If T_SolX = "ID" Then
            Cons = Cons & ")"
        Else
            Cons = Cons & " AND ((Orden.letra)='1'))"
        End If
        Cons = Cons & " ORDER BY Orden.ano DESC, Orden.Numero DESC, Orden.letra DESC" 'And (Abierta = 1)

        Using connection As New SqlConnection(cadenaconexion)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                cboNumOrden.Items.Clear()


                Do While reader.Read()
                    cboNumOrden.Items.Add(reader.GetValue(1) & "-" & reader.GetValue(0))
                Loop
            End If

            command = Nothing
            reader.Close()
            reader = Nothing

        End Using

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboNumOrden.SelectedIndexChanged
        Carga_Checklist(cboTipoOrden.SelectedItem, cboNumOrden.SelectedItem)
        planeacionSCI(cboNumOrden.SelectedItem)
    End Sub
    Public Sub planeacionSCI(Num_OF As String)
        DataGridView1.Item(1, 0).Value = ChekRayas.Items.Count

        Dim Cons As String = "SELECT Vista_ActaSeg_Completo.M2_cotizados, Orden.m2, Orden.Nu_Piezas, Orden.Muros, Orden.UML, Orden.Losa, Orden.Culata, Orden.Escalera, Orden.Punto_Fijo, Orden.Otros " +
                "FROM     Orden WITH (nolock) LEFT OUTER JOIN " +
                  "Vista_ActaSeg_Completo ON Orden.Id_Of_P = Vista_ActaSeg_Completo.idOfp " +
                "WHERE  (RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) AND (Orden.Muros = 1) AND (NOT (Orden.m2 IS NULL)) OR " +
                  "(RTrim(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) AND (Orden.UML = 1) AND (NOT (Orden.m2 IS NULL)) OR " +
                  "(RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) AND (Orden.Losa = 1) AND (NOT (Orden.m2 IS NULL)) OR " +
                  "(RTrim(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) AND (Orden.Culata = 1) AND (NOT (Orden.m2 IS NULL)) OR " +
                  "(RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) AND (Orden.Escalera = 1) AND (NOT (Orden.m2 IS NULL)) OR " +
                  "(RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) AND (Orden.Punto_Fijo = 1) AND (NOT (Orden.m2 IS NULL)) OR " +
                  "(RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) AND (Orden.Otros = 1) AND (NOT (Orden.m2 IS NULL)) " +
                "ORDER BY Orden.Id_Ofa"
        Dim rayasalum As Integer = 0
        Dim rayasacc As Integer = 0
        Dim metroscot As Double = 0
        Dim metrosfab As Double = 0
        Dim nuformaletas As Integer = 0
        Dim m2faltante As Double = 0
        Dim porcfaltante As Integer = 0
        Dim toneladasacc As Double = 0
        Using connection As New SqlConnection(cadenaconexion)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                Do While reader.Read()
                    rayasalum = rayasalum + 1

                    If metroscot = 0 Then
                        If IsDBNull(reader.GetValue(0)) Then
                        Else
                            metroscot = Math.Round(reader.GetValue(0), 2)
                        End If

                    End If
                    If IsDBNull(reader.GetValue(1)) Then
                    Else
                        metrosfab = metrosfab + Math.Round(reader.GetValue(1), 2)
                    End If
                    If IsDBNull(reader.GetValue(2)) Then
                    Else
                        nuformaletas = nuformaletas + reader.GetValue(2)
                    End If
                Loop
            End If

            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
        m2faltante = Math.Round(metroscot - metrosfab, 2)
        If m2faltante < 0 Then
            m2faltante = 0
        End If
        If metroscot = 0 Then
            porcfaltante = 0
        Else
            porcfaltante = 100 - (metrosfab * 100 / metroscot)
        End If


        Cons = "SELECT Orden.Ofa " +
                "FROM     Orden WITH (nolock) INNER JOIN " +
                  "Of_Accesorios WITH (nolock) ON Orden.Id_Ofa = Of_Accesorios.Id_Ofa " +
                "WHERE  (RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0) " +
                "GROUP BY Orden.Ofa, Orden.Id_Ofa " +
                "ORDER BY Orden.Id_Ofa"
        Using connection As New SqlConnection(cadenaconexion)
            Dim Command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = Command.ExecuteReader()

            If reader.HasRows Then
                Do While reader.Read()
                    rayasacc = rayasacc + 1
                Loop
            End If
            Command = Nothing
            reader.Close()
            reader = Nothing
        End Using
        Cons = "SELECT SUM(Of_Accesorios.Peso_Estimado * Of_Accesorios.Cant_Req) AS pesotot " +
                "FROM     Orden WITH (nolock) INNER JOIN " +
                  "Of_Accesorios WITH (nolock) ON Orden.Id_Ofa = Of_Accesorios.Id_Ofa " +
                "WHERE  (RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Orden.Anulada = 0)"
        Using connection As New SqlConnection(cadenaconexion)
            Dim Command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = Command.ExecuteReader()

            If reader.HasRows Then
                Do While reader.Read()
                    If IsDBNull(reader.GetValue(0)) Then
                        toneladasacc = 0
                    Else
                        toneladasacc = Math.Round(reader.GetValue(0) / 1000, 2)
                    End If

                Loop
            End If
            Command = Nothing
            reader.Close()
            reader = Nothing
        End Using

        DataGridView1.Item(1, 0).Value = rayasalum
        DataGridView1.Item(1, 1).Value = metroscot
        DataGridView1.Item(1, 2).Value = metrosfab
        DataGridView1.Item(1, 3).Value = nuformaletas
        DataGridView1.Item(1, 4).Value = m2faltante
        DataGridView1.Item(1, 5).Value = porcfaltante & "%"
        DataGridView1.Item(1, 6).Value = rayasacc
        DataGridView1.Item(1, 7).Value = toneladasacc
    End Sub
    Public Sub Carga_Checklist(ByRef T_SolX As String, ByRef Num_OF As String)
        Dim Cons As String = "SELECT Orden.letra AS Raya, Orden.Id_Ofa FROM Orden WITH (nolock) " +
            "WHERE (Orden.Tipo_Of = '" & T_SolX & "') AND (RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) LIKE '" & Num_OF & "%') AND (Anulada = 0)"
        If T_SolX = "ID" Then
        Else
            Cons = Cons & " AND (NOT (letra LIKE '%M'))"
        End If
        If CheckBox4.Checked = False Then
            Cons = Cons & "AND (Escalera = 0) "
        End If
        Cons = Cons & "ORDER BY RTRIM(Orden.Numero) + '-' + RTRIM(Orden.ano) DESC, Raya"
        '(Orden.Abierta = 0) AND 
        Using connection As New SqlConnection(cadenaconexion)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                ChekRayas.Items.Clear()
                'CheckedListBox1.Items.Add("Form")
                'CheckedListBox1.Items.Add("Acc")
                Do While reader.Read()
                    If reader.GetValue(0).ToString = "" And T_SolX = "ID" Then
                        ChekRayas.Items.Add(1)
                        ChekRayas.Height = 4 + (ChekRayas.Items.Count * 17)
                        Me.Height = Math.Max(ChekRayas.Height, 299) + 95 '17x5+10
                        ChekRayas.Items.Add(51)
                        ChekRayas.Height = 4 + (ChekRayas.Items.Count * 17)
                        Me.Height = Math.Max(ChekRayas.Height, 299) + 95 '17x5+10
                    Else
                        ChekRayas.Items.Add(reader.GetValue(0))
                        ChekRayas.Height = 4 + (ChekRayas.Items.Count * 17)
                        Me.Height = Math.Max(ChekRayas.Height, 299) + 95 '17x5+10
                    End If
                Loop
            End If

            command = Nothing
            reader.Close()
            reader = Nothing

        End Using
    End Sub

    Private Function raya(ByVal checkbox As String) As Double
        Try
            raya = CDbl(checkbox)
        Catch ex As Exception
            Dim texto As String = checkbox
            Dim cadena() As Char = texto.ToCharArray()
            raya = CDbl(CStr(cadena(0)))
        End Try
    End Function
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckSoloFM.CheckedChanged
        If CheckSoloFM.Checked = True Then
            For i = 0 To ChekRayas.Items.Count - 1
                If raya(ChekRayas.Items.Item(i)) < 51 Then
                    ChekRayas.SetItemChecked(i, True)
                End If
            Next
        Else
            For i = 0 To ChekRayas.Items.Count - 1
                If raya(ChekRayas.Items.Item(i)) < 51 Then
                    ChekRayas.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckSoloAcc.CheckedChanged
        If CheckSoloAcc.Checked = True Then
            For i = 0 To ChekRayas.Items.Count - 1
                If raya(ChekRayas.Items.Item(i)) >= 51 Then
                    ChekRayas.SetItemChecked(i, True)
                End If
            Next
        Else
            For i = 0 To ChekRayas.Items.Count - 1
                If raya(ChekRayas.Items.Item(i)) >= 51 Then
                    ChekRayas.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnDescargarListado.Click

        If cboNumOrden.Text Is Nothing Then
            MessageBox.Show("Seleccione una Orden para generar el Listado")
            Return
        End If

        combinar = False
        If checkStockERP.Checked = True Then
            descargalistaERP(False)
        ElseIf checkCliente.Checked = True Then
            descargarListaCliente(CheckSoloFM.Checked, CheckSoloAcc.Checked)
        Else
            descargalista(False)
        End If
    End Sub
    Sub descargalistaERP(ByVal ADAPO As Boolean)
        Dim lista(,) As String
        ReDim lista(18, 0)
        lista(0, UBound(lista, 2)) = "CANT"
        lista(1, UBound(lista, 2)) = "DESCRIPCION"
        lista(2, UBound(lista, 2)) = "ANCHO"
        lista(3, UBound(lista, 2)) = "ALTO"
        lista(4, UBound(lista, 2)) = "ALTO2"
        lista(5, UBound(lista, 2)) = "ANCHO2"
        lista(6, UBound(lista, 2)) = "PLANO"
        lista(7, UBound(lista, 2)) = "DESC_AUX"
        lista(8, UBound(lista, 2)) = "FAMILIA"
        lista(9, UBound(lista, 2)) = "AREA_ITEM"
        lista(10, UBound(lista, 2)) = "RAYA"
        lista(11, UBound(lista, 2)) = "ITEM"
        lista(12, UBound(lista, 2)) = "NUMERO_MEMO"
        lista(13, UBound(lista, 2)) = "TIPO_DEL_MEMO"
        lista(14, UBound(lista, 2)) = "DESCRIPCION_DEL_MEMORANDO"
        lista(15, UBound(lista, 2)) = "REF_ORDEN"
        lista(16, UBound(lista, 2)) = "OBSERVACIONES"
        ReDim Preserve lista(18, UBound(lista, 2) + 1)
        'Dim rsiferp2 = New ADODB.Recordset



        'Consulta Resumida. No genera TimeOut
        Dim Slrs10 = "SELECT " +
                     "T120_MC_ITEMS.F120_DESCRIPCION, " +
                     "BI_T400.F_CANT_DISPONIBLE_1 " +
                     "FROM " +
                     "T120_MC_ITEMS INNER JOIN " +
                     "BI_T400 " +
                     "On BI_T400.F_ROWID_ITEM = T120_MC_ITEMS.F120_ROWID " +
                     "WHERE " +
                     "BI_T400.F_ID_CIA = 6 AND " +
                     "BI_T400.F_ID_BODEGA = 'B08' AND " + '"BI_T400.F_ID_BODEGA = 'B09' And " 
                     "BI_T400.F_PARAMETRO_BIABLE = '6' AND " +
                     "BI_T400.F_ID_UBICACION_AUX Not Like '%CHATARRA%' AND " +
                     "BI_T400.F_CANT_DISPONIBLE_1 > 0 "
        'Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        Dim cadenaconexionERP = CreaRQI.MyCommands.AbreDB_ERP

        'rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablaItems(,) As Object
        ReDim Preserve tablaItems(1, 0)

        Try
            Using connection As New SqlConnection(cadenaconexionERP)
                Dim command As SqlCommand = New SqlCommand(Slrs10, connection)
                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    While reader.Read()
                        tablaItems(0, tablaItems.GetUpperBound(1)) = reader.GetValue(0)
                        tablaItems(1, tablaItems.GetUpperBound(1)) = reader.GetValue(1)
                        ReDim Preserve tablaItems(1, tablaItems.GetUpperBound(1) + 1)
                    End While

                End If
            End Using

        Catch ex As Exception

            MessageBox.Show("No fue Posible Establecer Conexion Con El ERP. Intente nuevamente con la Aplicacion Ejecutada desde el Servidor")
            Return

        End Try

        'tablaItems = rsiferp2.GetRows
        For i = 0 To tablaItems.GetUpperBound(1)
            Dim item = Split(tablaItems(0, i), " "c)
            'try catch para evitar detenerse cuando hay items sin el formato adecuado -nombre medida1 X medida2
            Dim t, dim3, dim4, angular
            Try
                t = CLng(item(1))
                t = CLng(item(3))
            Catch ex As Exception
                t = 640354 'numero que nunca saldria
            End Try
            angular = False
            dim3 = ""
            dim4 = ""
            If item(0) = "FMA" Or item(0) = "#FMA" Or item(0) = "CPA" Or item(0) = "#CPA" Or item(0) = "CPAT" Or item(0) = "#CPAT" Or item(0) = "EQLI" Or item(0) = "EQLE" Or item(0) = "#EQLI" Or item(0) = "#EQLE" Then
                angular = True
            End If
            Try
                If item(5) > 0 And angular = False Then
                    dim3 = item(5)
                    dim4 = ""
                ElseIf item(5) > 0 And angular = True Then
                    dim3 = ""
                    dim4 = item(5)
                End If
            Catch ex As Exception
                dim3 = ""
                dim4 = ""
            End Try
            If UBound(item, 1) < 3 Or t = 640354 Then

            Else
                lista(0, UBound(lista, 2)) = tablaItems(1, i)   'cantidad
                lista(1, UBound(lista, 2)) = item(0)            'nomenclatura
                lista(2, UBound(lista, 2)) = item(1)            'ancho1
                lista(3, UBound(lista, 2)) = item(3)            'alto1
                lista(4, UBound(lista, 2)) = dim3               'alto2
                lista(5, UBound(lista, 2)) = dim4               'ancho2
                lista(6, UBound(lista, 2)) = ""                 'Plano Especial
                lista(7, UBound(lista, 2)) = ""                 'observacion
                lista(8, UBound(lista, 2)) = ""                 'Familia
                lista(9, UBound(lista, 2)) = ""                 'Area
                lista(10, UBound(lista, 2)) = ""                'OF
                lista(11, UBound(lista, 2)) = ""                'Item
                lista(12, UBound(lista, 2)) = ""                'Memo
                lista(13, UBound(lista, 2)) = ""                'Tipo de cambio
                lista(14, UBound(lista, 2)) = ""                'Observacion Memo
                lista(15, UBound(lista, 2)) = ""                'ref orde
                lista(16, UBound(lista, 2)) = ""                'Observaciones
                ReDim Preserve lista(18, UBound(lista, 2) + 1)
            End If
        Next
        Do While lista(0, UBound(lista, 2)) = Nothing
            ReDim Preserve lista(18, UBound(lista, 2) - 1)
        Loop


        Dim saveFileDialog As New SaveFileDialog()
        Dim ExtensionFile As String = ".xls"
        saveFileDialog.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        saveFileDialog.Title = "Guardar " & "Lista Orden " & cboNumOrden.SelectedItem
        saveFileDialog.FileName = "Lista Stock " & ExtensionFile


        If saveFileDialog.ShowDialog() = DialogResult.OK Then

            Dim filepath = saveFileDialog.FileName
            CreaRQI.MyCommands.escribe(filepath, lista, True)

            If File.Exists(filepath) Then

                MessageBox.Show("Archivo Guardado Con Exito")

            End If

        End If


    End Sub

    Sub descargalista(ByVal ADAPO As Boolean)

        Dim lista(,) As String
        Dim idofaP As String

        If combinar = True And primercombi = False Then
            lista = listacombi
        ElseIf combinar = False And primercombi = False Then
            lista = listacombi
        Else
            primercombi = False
            ReDim lista(18, 0)
            lista(0, UBound(lista, 2)) = "CANT"
            lista(1, UBound(lista, 2)) = "DESCRIPCION"
            lista(2, UBound(lista, 2)) = "ANCHO"
            lista(3, UBound(lista, 2)) = "ALTO"
            lista(4, UBound(lista, 2)) = "ALTO2"
            lista(5, UBound(lista, 2)) = "ANCHO2"
            lista(6, UBound(lista, 2)) = "PLANO"
            lista(7, UBound(lista, 2)) = "DESC_AUX"
            lista(8, UBound(lista, 2)) = "FAMILIA"
            lista(9, UBound(lista, 2)) = "AREA_ITEM"
            lista(10, UBound(lista, 2)) = "RAYA"
            lista(11, UBound(lista, 2)) = "ITEM"
            lista(12, UBound(lista, 2)) = "NUMERO_MEMO"
            lista(13, UBound(lista, 2)) = "TIPO_DEL_MEMO"
            lista(14, UBound(lista, 2)) = "DESCRIPCION_DEL_MEMORANDO"
            lista(15, UBound(lista, 2)) = "REF_ORDEN"
            lista(16, UBound(lista, 2)) = "DIMENSIONES"
            lista(17, UBound(lista, 2)) = "PESOUNI"
            lista(18, UBound(lista, 2)) = "PESOTOT"
            ReDim Preserve lista(18, UBound(lista, 2) + 1)
        End If

        If lista Is Nothing Then
            primercombi = False
            ReDim lista(18, 0)
            lista(0, UBound(lista, 2)) = "CANT"
            lista(1, UBound(lista, 2)) = "DESCRIPCION"
            lista(2, UBound(lista, 2)) = "ANCHO"
            lista(3, UBound(lista, 2)) = "ALTO"
            lista(4, UBound(lista, 2)) = "ALTO2"
            lista(5, UBound(lista, 2)) = "ANCHO2"
            lista(6, UBound(lista, 2)) = "PLANO"
            lista(7, UBound(lista, 2)) = "DESC_AUX"
            lista(8, UBound(lista, 2)) = "FAMILIA"
            lista(9, UBound(lista, 2)) = "AREA_ITEM"
            lista(10, UBound(lista, 2)) = "RAYA"
            lista(11, UBound(lista, 2)) = "ITEM"
            lista(12, UBound(lista, 2)) = "NUMERO_MEMO"
            lista(13, UBound(lista, 2)) = "TIPO_DEL_MEMO"
            lista(14, UBound(lista, 2)) = "DESCRIPCION_DEL_MEMORANDO"
            lista(15, UBound(lista, 2)) = "REF_ORDEN"
            lista(16, UBound(lista, 2)) = "DIMENSIONES"
            lista(17, UBound(lista, 2)) = "PESOUNI"
            lista(18, UBound(lista, 2)) = "PESOTOT"
            ReDim Preserve lista(18, UBound(lista, 2) + 1)
        End If


        For i = 0 To ChekRayas.Items.Count - 1
            Dim raya = ChekRayas.Items.Item(i)
            If ChekRayas.GetItemChecked(i) And ChekRayas.Items.Item(i) < 51 Then
                Dim Cons As String = "Select ISNULL(Memo_Det.Memo_DetCantPFin, Saldos.cant) As Cant, Saldos.Tipo, Saldos.Anc1, Saldos.alto1, Saldos.Alto2, Saldos.Anc2, " +
                "Saldos.Plano_esp, IIF(saldos.Observacion = '.', '', Saldos.Observacion) AS Observacion, Saldos.Grupo, Saldos.Area, Orden.Ofa, Saldos.Item, Saldos.Ult_Memo, Memo_Det.Memo_DetOperacion, Memo_Det.Memo_Obs, Orden.Id_Of_P, Orden.Ref, Formaleta_Lib_Inv.Formaleta_Lib_Inv_Peso AS PesoUni, Formaleta_Lib_Inv.Formaleta_Lib_Inv_Peso * Saldos.Cant_Final_Req AS PesoTotal " +
                "FROM Explo_Saldos WITH (nolock) INNER JOIN Formaleta_Lib_Inv WITH (nolock) ON Explo_Saldos.Explo_Lib_Id = Formaleta_Lib_Inv.Formaleta_Lib_Inv_Id RIGHT OUTER JOIN " +
                "Saldos WITH (nolock) INNER JOIN Orden WITH (nolock) ON Saldos.Id_Ofa = Orden.Id_Ofa ON Explo_Saldos.Saldos_Id = Saldos.Identificador LEFT OUTER JOIN " +
                "Memos WITH (nolock) INNER JOIN Memo_Det WITH (nolock) ON Memos.Id_Memo = Memo_Det.Id_MemoId ON Saldos.Ult_Memo = Memos.Memo_No AND Saldos.Identificador = Memo_Det.Memo_DetIdSaldosOri " +
                "WHERE (Orden.Ofa LIKE '" & cboNumOrden.SelectedItem & "-" & ChekRayas.Items.Item(i) & "') AND (Saldos.Anula = 0) OR " +
                "(Orden.Ofa LIKE '" & cboNumOrden.SelectedItem & "-" & ChekRayas.Items.Item(i) & "M') AND (Saldos.Anula = 0) " +
                "ORDER BY Saldos.Grupo, Saldos.Item"

                '"SELECT Saldos.cant, Saldos.Tipo, Saldos.Anc1, Saldos.alto1, Saldos.Alto2, Saldos.Anc2, Saldos.Observacion, Saldos.Grupo " + _
                '"FROM  Saldos INNER JOIN " + _
                '"Orden ON Saldos.Id_Ofa = Orden.Id_Ofa " + _
                '"WHERE (Orden.Ofa = '" & ComboBox2.SelectedItem & "-" & CheckedListBox1.Items.Item(i) & "')"

                Using connection As New SqlConnection(cadenaconexion)
                    Dim command As SqlCommand = New SqlCommand(Cons, connection)
                    connection.Open()
                    Dim reader As SqlDataReader = command.ExecuteReader()

                    If reader.HasRows Then
                        Do While reader.Read()
                            If idofaP = Nothing Then
                                idofaP = reader.GetValue(15)
                            End If
                            'If lista(i, 0) = "" Then
                            'lista(i, 0) = reader.GetValue(0)
                            'lista(i, 1) = reader.GetValue(1)
                            'lista(i, 2) = reader.GetValue(2)
                            'lista(i, 3) = reader.GetValue(3)
                            'lista(i, 4) = reader.GetValue(4)
                            'lista(i, 5) = reader.GetValue(5)
                            'lista(i, 6) = reader.GetValue(6)
                            'lista(i, 7) = reader.GetValue(7)
                            'lista(i, 8) = reader.GetValue(8)
                            'Else
                            Dim desaux, dim3, dim4, memo, TipoCambio, Obser, Ref, Plano, PesoUni, PesoTot As String

                            If reader.GetValue(4) = "0" Then
                                dim3 = ""
                            Else
                                dim3 = reader.GetValue(4)
                            End If
                            If reader.GetValue(5) = "0" Then
                                dim4 = ""
                            Else
                                dim4 = reader.GetValue(5)
                            End If
                            If reader.Item(6) = True Then
                                Plano = "P"
                            Else
                                Plano = ""
                            End If
                            If reader.GetValue(7) = " " Then
                                desaux = ""
                            Else
                                desaux = reader.GetValue(7)
                            End If


                            If IsDBNull(reader.GetValue(12)) Then
                                memo = ""
                            Else
                                memo = "memo " & reader.GetValue(12)
                            End If
                            If IsDBNull(reader.GetValue(13)) Then
                                TipoCambio = ""
                            Else
                                TipoCambio = reader.GetValue(13)
                            End If
                            If IsDBNull(reader.GetValue(14)) Then
                                Obser = ""
                            Else
                                Obser = reader.GetValue(14)
                            End If
                            If IsDBNull(reader.GetValue(16)) Then
                                Ref = ""
                            Else
                                Ref = reader.GetValue(16)
                            End If
                            If IsDBNull(reader.GetValue(17)) Then
                                PesoUni = 0
                            Else
                                PesoUni = Math.Round(reader.GetValue(17), 1)
                            End If
                            If IsDBNull(reader.GetValue(18)) Then
                                PesoTot = 0
                            Else
                                PesoTot = Math.Round(reader.GetValue(18), 1)
                            End If

                            Dim concatenado = reader.GetValue(1) & "-" & reader.GetValue(2) & "-" & reader.GetValue(3) & "-" & dim3 & "-" & dim4 & "-" & desaux & "-" & reader.GetValue(8) & "-" & reader.GetValue(10) & "-" & reader.GetValue(11) & "-" & memo & "-" & TipoCambio & "-" & Obser
                            Dim sumar As Boolean = False
                            Dim k As Integer

                            For j = 0 To UBound(lista, 2)
                                Dim prueba = lista(1, j) & "-" & lista(2, j) & "-" & lista(3, j) & "-" & lista(4, j) & "-" & lista(5, j) & "-" & lista(7, j) & "-" & lista(8, j) & "-" & lista(10, j) & "-" & lista(11, j) & "-" & lista(12, j) & "-" & lista(13, j) & "-" & lista(14, j)
                                If concatenado = prueba Then
                                    sumar = True
                                    k = j
                                    Exit For
                                End If
                            Next

                            If sumar = True Then
                                lista(0, k) = CDbl(lista(0, k)) + CDbl(reader.GetValue(0))
                            Else
                                lista(0, UBound(lista, 2)) = reader.GetValue(0) 'cantidad
                                lista(1, UBound(lista, 2)) = reader.GetValue(1) 'nomenclatura
                                lista(2, UBound(lista, 2)) = reader.GetValue(2) 'ancho1
                                lista(3, UBound(lista, 2)) = reader.GetValue(3) 'alto1
                                lista(4, UBound(lista, 2)) = dim3               'alto2
                                lista(5, UBound(lista, 2)) = dim4               'ancho2
                                lista(6, UBound(lista, 2)) = Plano              'Plano Especial
                                lista(7, UBound(lista, 2)) = desaux             'observacion
                                lista(8, UBound(lista, 2)) = reader.GetValue(8) 'Familia
                                lista(9, UBound(lista, 2)) = reader.GetValue(9) 'Area
                                lista(10, UBound(lista, 2)) = reader.GetValue(10) 'OF
                                lista(11, UBound(lista, 2)) = reader.GetValue(11) 'Item
                                lista(12, UBound(lista, 2)) = memo              'Memo
                                lista(13, UBound(lista, 2)) = TipoCambio        'Tipo de cambio
                                lista(14, UBound(lista, 2)) = Obser             'Observacion Memo
                                lista(15, UBound(lista, 2)) = Ref               'Ref Orden
                                lista(16, UBound(lista, 2)) = reader.GetValue(2) & "x" & reader.GetValue(3) & "x" & If(dim3 = "", 0, dim3) & "x" & If(dim4 = "", 0, dim4)
                                lista(17, UBound(lista, 2)) = PesoUni
                                lista(18, UBound(lista, 2)) = PesoTot
                                ReDim Preserve lista(18, UBound(lista, 2) + 1)
                            End If
                            'End If
                            'If lista(i, 8) = "" Then

                            'End If
                            'CheckedListBox1.Items.Add(reader.GetValue(0))
                            'CheckedListBox1.Height = 4 + (CheckedListBox1.Items.Count * 17)
                            'Me.Height = Math.Max(CheckedListBox1.Height, 55) + 36 + 34
                        Loop
                        'Dim raya = lista(10, UBound(lista, 2) - 1)
                    End If

                    command = Nothing
                    reader.Close()
                    reader = Nothing

                End Using

            End If

        Next

        For i = 0 To ChekRayas.Items.Count - 1
            Dim raya = ChekRayas.Items.Item(i)
            If ChekRayas.GetItemChecked(i) And ChekRayas.Items.Item(i) > 50 Then
                'Anterior - Antiguo
                'Dim Cons As String = "SELECT DISTINCT ISNULL(Memo_Acc_Det.Memo_Acc_Cant_Fin, Of_Accesorios.Cant_Req) AS Cant_Req, CASE WHEN Of_Accesorios.Nomenclatura != '' THEN Of_Accesorios.Nomenclatura " +
                '"WHEN Accesorios_Codigos.Nomenclatura !='' THEN Accesorios_Codigos.Nomenclatura ELSE item_planta.descripcion END AS Nomenclatura," +
                '"Of_Accesorios.Des_Aux, Of_Accesorios.Dim1, Of_Accesorios.Dim2, Of_Accesorios.Dim3, Of_Accesorios.Dim4, Of_Accesorios.Dim5, " +
                '"Of_Accesorios.Dim6, Orden.Ofa, Of_Accesorios.No_Item, Orden.Id_Of_P, Orden.Id_Ofa, Of_Accesorios.Ult_Memo, Memo_Acc_Det.Memo_Acc_Proceso," +
                '"Memo_Acc_Det.Memo_Acc_Obs, Of_Accesorios.Peso_Estimado AS PesoUni, Of_Accesorios.Peso_Estimado * Of_Accesorios.Cant_Req AS PesoTotal " +
                '"FROM  Of_Accesorios WITH (nolock) INNER JOIN " +
                '"Orden WITH (nolock) ON Of_Accesorios.Id_Ofa = Orden.Id_Ofa INNER JOIN " +
                '"item_planta WITH (nolock) ON Orden.planta_id = item_planta.planta_id AND Of_Accesorios.Id_UnoE = item_planta.cod_erp LEFT OUTER JOIN " +
                '"Memo_Acc_Det WITH (nolock) ON Of_Accesorios.Id_Orden_Acce = Memo_Acc_Det.Memo_Acc_OfAccId LEFT OUTER JOIN " +
                '"Accesorios_Codigos WITH (nolock) ON Of_Accesorios.Id_UnoE = Accesorios_Codigos.Id_UnoE AND Orden.planta_id = Accesorios_Codigos.planta_id " +
                '"WHERE (Orden.Ofa LIKE '" & ComboBox2.SelectedItem & "-" & CheckedListBox1.Items.Item(i) & "%') AND (Of_Accesorios.Anula = 0) " +
                '"ORDER BY Orden.Ofa, Of_Accesorios.No_Item"

                'Actualizado
                Dim Cons As String = "SELECT DISTINCT ISNULL(Memo_Acc_Det.Memo_Acc_Cant_Fin, Of_Accesorios.Cant_Req) AS Cant_Req, " +
                "CASE " +
                "WHEN Of_Accesorios.Nomenclatura != '' THEN Of_Accesorios.Nomenclatura " + 'Se ajusta esta linea por los negativos, invirtiendo el orden por la linea siguiente
                "WHEN Accesorios_Codigos.Nomenclatura !='' THEN Accesorios_Codigos.Nomenclatura " +
                "ELSE item_planta.descripcion END AS Nomenclatura, " +
                "Of_Accesorios.Des_Aux, " +
                "Of_Accesorios.Dim1, " +
                "Of_Accesorios.Dim2, " +
                "Of_Accesorios.Dim3, " +
                "Of_Accesorios.Dim4, " +
                "Of_Accesorios.Dim5, " +
                "Of_Accesorios.Dim6, " +
                "Orden.Ofa, " +
                "Of_Accesorios.No_Item, " +
                "Orden.Id_Of_P, " +
                "Orden.Id_Ofa, " +
                "Of_Accesorios.Ult_Memo, " +
                "Memo_Acc_Det.Memo_Acc_Proceso, " +
                "Memo_Acc_Det.Memo_Acc_Obs, " +
                "Of_Accesorios.Peso_Estimado AS PesoUni, " +
                "Of_Accesorios.Peso_Estimado * Of_Accesorios.Cant_Req AS PesoTotal " +
                "FROM  Of_Accesorios WITH (nolock) INNER JOIN Orden WITH (nolock) " +
                "ON Of_Accesorios.Id_Ofa = Orden.Id_Ofa " +
                "INNER JOIN item_planta WITH (nolock) " +
                "ON Orden.planta_id = item_planta.planta_id AND Of_Accesorios.Id_UnoE = item_planta.cod_erp " +
                "LEFT OUTER JOIN Memo_Acc_Det WITH (nolock) ON Of_Accesorios.Id_Orden_Acce = Memo_Acc_Det.Memo_Acc_OfAccId " +
                "LEFT JOIN Accesorios_Codigos WITH(nolock) ON Accesorios_Codigos.Codigos_Id = Of_Accesorios.AccesoriosCodigosId AND Orden.planta_id = Accesorios_Codigos.planta_id " +
                "WHERE (Orden.Ofa LIKE '" & cboNumOrden.SelectedItem & "-" & ChekRayas.Items.Item(i) & "%') AND (Of_Accesorios.Anula = 0) " +
                "ORDER BY Orden.Ofa, Of_Accesorios.No_Item"

                Using connection As New SqlConnection(cadenaconexion)
                    Dim command As SqlCommand = New SqlCommand(Cons, connection)
                    connection.Open()
                    Dim reader As SqlDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        Do While reader.Read()
                            If idofaP = Nothing Then
                                idofaP = reader.GetValue(11)
                            End If
                            Dim nomenclatura, desaux, dim1, dim2, dim3, dim4, dim5, dim6, memo, TipoCambio, Obser, Plano, PesoUni, PesoTot As String
                            nomenclatura = RemoveXtraSpaces(reader.GetValue(1))
                            If reader.GetValue(3) = "0" Then
                                dim1 = ""
                            Else
                                dim1 = reader.GetValue(3)
                            End If
                            If reader.GetValue(4) = "0" Then
                                dim2 = ""
                            Else
                                dim2 = reader.GetValue(4)
                            End If
                            If reader.GetValue(5) = "0" Then
                                dim3 = ""
                            Else
                                dim3 = reader.GetValue(5)
                            End If
                            If reader.GetValue(6) = "0" Then
                                dim4 = ""
                            Else
                                dim4 = reader.GetValue(6)
                            End If
                            If reader.GetValue(7) = "0" Then
                                dim5 = ""
                            Else
                                dim5 = reader.GetValue(7)
                            End If
                            If reader.GetValue(8) = "0" Then
                                dim6 = ""
                            Else
                                dim6 = reader.GetValue(8)
                            End If
                            If IsDBNull(reader.GetValue(13)) Then
                                memo = ""
                            Else
                                memo = "memo " & reader.GetValue(13)
                            End If
                            If IsDBNull(reader.GetValue(14)) Then
                                TipoCambio = ""
                            Else
                                TipoCambio = reader.GetValue(14)
                            End If
                            If IsDBNull(reader.GetValue(15)) Then
                                Obser = ""
                            Else
                                Obser = reader.GetValue(15)
                            End If
                            If IsDBNull(reader.GetValue(16)) Then
                                PesoUni = 0
                            Else
                                PesoUni = Math.Round(reader.GetValue(16), 1)
                            End If
                            If IsDBNull(reader.GetValue(17)) Then
                                PesoTot = 0
                            Else
                                PesoTot = Math.Round(reader.GetValue(17), 1)
                            End If
                            If IsDBNull(reader.GetValue(2)) Then
                                desaux = ""
                            ElseIf reader.GetValue(2) = " " Then
                                desaux = ""
                            Else
                                desaux = reader.GetValue(2)
                            End If

                            lista(0, UBound(lista, 2)) = reader.GetValue(0)     'cantidad
                            lista(1, UBound(lista, 2)) = nomenclatura           'nomenclatura
                            lista(2, UBound(lista, 2)) = desaux                 'aux
                            lista(3, UBound(lista, 2)) = dim1                   'alto1
                            lista(4, UBound(lista, 2)) = dim2                   'alto2
                            lista(5, UBound(lista, 2)) = dim3                   'ancho2
                            lista(6, UBound(lista, 2)) = dim4                   'Plano Especial
                            lista(7, UBound(lista, 2)) = dim5                   'observacion
                            lista(8, UBound(lista, 2)) = dim6                   'Familia
                            lista(9, UBound(lista, 2)) = ""                     'Area
                            lista(10, UBound(lista, 2)) = reader.GetValue(9)    'OF
                            lista(11, UBound(lista, 2)) = reader.GetValue(10)   'Item
                            lista(12, UBound(lista, 2)) = memo                  'Memo
                            lista(13, UBound(lista, 2)) = TipoCambio            'Tipo de cambio
                            lista(14, UBound(lista, 2)) = Obser                 'Observacion Memo
                            lista(15, UBound(lista, 2)) = ""                    'Ref Orden
                            lista(16, UBound(lista, 2)) = If(dim1 = "", 0, dim1) & "x" & If(dim2 = "", 0, dim2) & "x" & If(dim3 = "", 0, dim3) & "x" & If(dim4 = "", 0, dim4) & "x" & If(dim5 = "", 0, dim5) & "x" & If(dim6 = "", 0, dim6)
                            lista(17, UBound(lista, 2)) = PesoUni
                            lista(18, UBound(lista, 2)) = PesoTot
                            ReDim Preserve lista(18, UBound(lista, 2) + 1)

                        Loop
                        'Dim raya = lista(10, UBound(lista, 2) - 1)
                    End If
                End Using
            End If
        Next

        If cboTipoOrden.SelectedItem = "ID" And idofaP = Nothing Then
            Dim cons = "SELECT DISTINCT Orden.Id_Of_P FROM Orden WITH (nolock) INNER JOIN Orden_Seg WITH (nolock) ON Orden.Id_Of_P = Orden_Seg.Id_Ofa WHERE (RTRIM(Orden_Seg.Num_Of) + '-' + RTRIM(Orden_Seg.Ano_Of) = '" & cboNumOrden.SelectedItem & "')"
            Using connection As New SqlConnection(cadenaconexion)
                Dim command As SqlCommand = New SqlCommand(cons, connection)
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()
                If reader.HasRows Then
                    Do While reader.Read()
                        idofaP = reader.GetValue(0)
                    Loop
                End If
            End Using
        End If

        If idofaP <> Nothing Then
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, cadenaconexion, "OG")
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, cadenaconexion, "OM")
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, cadenaconexion, "RC")
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, cadenaconexion, "ID")
            CreaRQI.MyCommands.buscagarantiaacc(lista, cboNumOrden.SelectedItem, cadenaconexion)
        End If

        If combinar = True Then
            listacombi = lista
            MsgBox("Seleccione un listado nuevo para combinar, si la proxima elección es el ultimo listado elija la opcion 'Descargar Listado' para que no pida otro adicional")
            Me.Show()
        Else

            Do While lista(0, UBound(lista, 2)) = Nothing
                ReDim Preserve lista(18, UBound(lista, 2) - 1)
            Loop

            Dim filename = CreaRQI.MyCommands.RutaPlanoAcad

            If filename = Nothing Then

                Dim saveFileDialog As New SaveFileDialog()
                Dim ExtensionFile As String = ".xls"
                saveFileDialog.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
                saveFileDialog.Title = "Guardar " & "Lista Orden " & cboNumOrden.SelectedItem
                saveFileDialog.FileName = "Lista Orden " & cboNumOrden.SelectedItem & ExtensionFile


                If saveFileDialog.ShowDialog() = DialogResult.OK Then

                    Dim filepath = saveFileDialog.FileName
                    CreaRQI.MyCommands.escribe(filepath, lista, True)

                    If File.Exists(filepath) Then

                        MessageBox.Show("Archivo Guardado Con Exito")

                    End If

                End If

            End If
        End If
    End Sub

    Public Sub descargarListaCliente(ByVal ListAlum As Boolean, ByVal ListAcc As Boolean)

        'Los parametros de la funcion dependen de los checks que hay en el formulario (CheckForm - CheckAcc)
        Dim NumOrden As String = cboNumOrden.Text
        Dim dt As DataTable
        Dim InitialDirectory As String

        Try

            InitialDirectory = CreaRQI.MyCommands.RutaPlanoAcad

        Catch ex As Exception

            InitialDirectory = Nothing

        End Try

        'Si ambos NO estan selccionados
        If ListAlum = False And ListAcc = False Then

            MessageBox.Show("Debe Seleccionar una Opcion para generar el listado, Aluminio o Accesorios")
            Return

        End If

        'Si ambos SI estan selccionados
        If ListAlum = True And ListAcc = True Then

            MessageBox.Show("Seleccione solo una Opcion para generar el listado, Aluminio o Accesorios")
            Return

        End If

        'Si solo Formaletas esta seleccionado
        If ListAlum = True And ListAcc = False Then

            dt = Ord.ConsultarListadoClienteAlum(NumOrden)

            If dt Is Nothing Then

                MessageBox.Show("No se encontraron Registros")
                Return

            End If

            Util.ProExportarExcelSimple(dt, "Listado Compilado Orden " & NumOrden, InitialDirectory)

        End If

        'Si solo Accesorios esta seleccionado
        If ListAlum = False And ListAcc = True Then

            dt = Ord.ConsultarListadoClienteACC(NumOrden)

            If dt Is Nothing Then

                MessageBox.Show("No se encontraron Registros")
                Return

            End If

            Util.ProExportarExcelSimple(dt, "Listado Compilado Orden " & NumOrden, InitialDirectory)

        End If

    End Sub

    Public Shared Function RemoveXtraSpaces(strVal As String) As String
        Dim iCount As Integer = 1
        Dim sTempstrVal As String

        sTempstrVal = ""

        For iCount = 1 To Len(strVal)
            sTempstrVal = sTempstrVal + Mid(strVal, iCount, 1).Trim
        Next

        RemoveXtraSpaces = sTempstrVal

        Return RemoveXtraSpaces

    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        combinar = False
        If checkStockERP.Checked = True Then
            descargalistaERP(True)
        Else
            descargalista(True)
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles checkStockERP.CheckedChanged
        If checkStockERP.Checked = True Then
            cboTipoOrden.Enabled = False
            cboNumOrden.Enabled = False
            CheckSoloFM.Enabled = False
            CheckSoloAcc.Enabled = False
            ChekRayas.Enabled = False
        Else
            cboTipoOrden.Enabled = True
            cboNumOrden.Enabled = True
            CheckSoloFM.Enabled = True
            CheckSoloAcc.Enabled = True
            ChekRayas.Enabled = True
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        CheckSoloFM.Checked = False
        CheckSoloAcc.Checked = False
        Carga_Checklist(cboTipoOrden.SelectedItem, cboNumOrden.SelectedItem)
    End Sub

    Private Sub CheckedListBox1_Itemcheck(sender As Object, e As EventArgs) Handles ChekRayas.ItemCheck

    End Sub

    Private Sub Frm_ListFormaletas_Load(sender As Object, e As EventArgs) Handles Me.Load

        llenacomboTipoOrden()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()
        DataGridView1.Rows.Add()

        DataGridView1.Item(0, 0).Value = "Cant Rayas Alum"
        DataGridView1.Item(1, 0).Value = "0"

        DataGridView1.Item(0, 1).Value = "m2 Cotizados"
        DataGridView1.Item(1, 1).Value = "0"
        DataGridView1.Rows.Item(1).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
        'Dim dgc1 = DataGridView1.Item(1, 1)
        'dgc1.Style.Font = New Drawing.Font(DataGridView1.DefaultCellStyle.Font.FontFamily, 12)

        DataGridView1.Item(0, 2).Value = "m2 Fabricados"
        DataGridView1.Item(1, 2).Value = "0"

        DataGridView1.Item(0, 3).Value = "Formaletas"
        DataGridView1.Item(1, 3).Value = "0"
        DataGridView1.Rows.Item(3).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray

        DataGridView1.Item(0, 4).Value = "m2 Faltantes"
        DataGridView1.Item(1, 4).Value = "0"

        DataGridView1.Item(0, 5).Value = "% m2 Faltantes"
        DataGridView1.Item(1, 5).Value = "0"
        DataGridView1.Rows.Item(5).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray

        DataGridView1.Item(0, 6).Value = "Cant Rayas Acc"
        DataGridView1.Item(1, 6).Value = "0"

        DataGridView1.Item(0, 7).Value = "Toneladas Acc"
        DataGridView1.Item(1, 7).Value = "0"
        DataGridView1.Rows.Item(7).DefaultCellStyle.BackColor = System.Drawing.Color.LightGray
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        combinar = True
        If checkStockERP.Checked = True Then
            descargalistaERP(False)
        Else
            descargalista(False)
        End If
    End Sub

    Sub llenacomboTipoOrden()
        Dim cons As String = "SELECT DISTINCT T_Sol_Tipo FROM Tipos_Sol WITH (nolock) WHERE (T_Sol_Activo = 1) AND (NOT (T_Sol_Tipo = 'OK'))"
        Using connection As New SqlConnection(cadenaconexion)
            Dim command As SqlCommand = New SqlCommand(cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                cboTipoOrden.Items.Clear() 'Combo ERP
                Do While reader.Read()
                    cboTipoOrden.Items.Add(reader.GetValue(0)) 'Combo ERP
                Loop
            End If
            command = Nothing
            reader.Close()
            reader = Nothing
        End Using
    End Sub

    Private Sub checkCliente_CheckedChanged(sender As Object, e As EventArgs) Handles checkCliente.CheckedChanged

        If ChekRayas.Enabled = True Then

            ChekRayas.Enabled = False
        Else
            ChekRayas.Enabled = True

        End If

    End Sub
End Class