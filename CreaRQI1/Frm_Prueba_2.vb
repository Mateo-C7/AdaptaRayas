Imports Microsoft.Office.Interop.Excel
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports System.Linq
Imports System.Data
Imports System.Collections.Generic
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.IO
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.ComponentModel.ListSortDirection
Imports Newtonsoft.Json
Imports System.Configuration
Imports CreaRQI.CapaDatos

Public Class Frm_Prueba_2
    Dim Conn As ConexionBD = New ConexionBD 'Instanciamos la clase de conexion
    Dim combinar As Boolean = False
    Dim primercombi As Boolean = True
    Dim listacombi(,) As String
    'Dim cadenaconexion = "data source = 172.21.224.130 ; persist security info=False;initial catalog = forsa ; user id=forsa; password=forsa2006 "
    Public Str_Con_Bd As String = ConexionBD.getStringConexion()
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Carga_Combo2(ComboBox1.SelectedItem)
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

        Using connection As New SqlConnection(Str_Con_Bd)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()
            Dim reader As SqlDataReader = command.ExecuteReader()
            If reader.HasRows Then
                ComboBox2.Items.Clear()
                Do While reader.Read()
                    ComboBox2.Items.Add(reader.GetValue(1) & "-" & reader.GetValue(0))
                Loop
            End If

            command = Nothing
            reader.Close()
            reader = Nothing

        End Using

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Carga_Checklist(ComboBox1.SelectedItem, ComboBox2.SelectedItem)
        planeacionSCI(ComboBox2.SelectedItem)
    End Sub
    Public Sub planeacionSCI(Num_OF As String)
        DataGridView1.Item(1, 0).Value = CheckedListBox1.Items.Count

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
        Using connection As New SqlConnection(Str_Con_Bd)

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
        Using connection As New SqlConnection(Str_Con_Bd)
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
        Using connection As New SqlConnection(Str_Con_Bd)
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
        Using connection As New SqlConnection(Str_Con_Bd)

            Dim command As SqlCommand = New SqlCommand(Cons, connection)

            connection.Open()

            Dim reader As SqlDataReader = command.ExecuteReader()

            If reader.HasRows Then
                CheckedListBox1.Items.Clear()
                'CheckedListBox1.Items.Add("Form")
                'CheckedListBox1.Items.Add("Acc")
                Do While reader.Read()
                    If reader.GetValue(0).ToString = "" And T_SolX = "ID" Then
                        CheckedListBox1.Items.Add(1)
                        CheckedListBox1.Height = 4 + (CheckedListBox1.Items.Count * 17)
                        Me.Height = Math.Max(CheckedListBox1.Height, 299) + 95 '17x5+10
                        CheckedListBox1.Items.Add(51)
                        CheckedListBox1.Height = 4 + (CheckedListBox1.Items.Count * 17)
                        Me.Height = Math.Max(CheckedListBox1.Height, 299) + 95 '17x5+10
                    Else
                        CheckedListBox1.Items.Add(reader.GetValue(0))
                        CheckedListBox1.Height = 4 + (CheckedListBox1.Items.Count * 17)
                        Me.Height = Math.Max(CheckedListBox1.Height, 299) + 95 '17x5+10
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
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            For i = 0 To CheckedListBox1.Items.Count - 1
                If raya(CheckedListBox1.Items.Item(i)) < 51 Then
                    CheckedListBox1.SetItemChecked(i, True)
                End If
            Next
        Else
            For i = 0 To CheckedListBox1.Items.Count - 1
                If raya(CheckedListBox1.Items.Item(i)) < 51 Then
                    CheckedListBox1.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            For i = 0 To CheckedListBox1.Items.Count - 1
                If raya(CheckedListBox1.Items.Item(i)) >= 51 Then
                    CheckedListBox1.SetItemChecked(i, True)
                End If
            Next
        Else
            For i = 0 To CheckedListBox1.Items.Count - 1
                If raya(CheckedListBox1.Items.Item(i)) >= 51 Then
                    CheckedListBox1.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        combinar = False
        If CheckBox3.Checked = True Then
            descargalistaERP(False)
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

        'Consulta Original para sacar el STOCK. Genera TimeOut
        'Dim Slrs10 = "Select " + 
        '                "T120_MC_ITEMS.F120_DESCRIPCION, " +
        '                "BI_T400.F_CANT_DISPONIBLE_1 " +
        '            "From " +
        '                "T120_MC_ITEMS Inner Join " +
        '                "BI_T400 " +
        '                    "On BI_T400.F_ROWID_ITEM = T120_MC_ITEMS.F120_ROWID Inner Join " +
        '                "T155_MC_UBICACION_AUXILIARES " +
        '                    "On T155_MC_UBICACION_AUXILIARES.F155_ID = BI_T400.F_ID_UBICACION_AUX " +
        '                "Inner Join " +
        '                "T150_MC_BODEGAS " +
        '                    "On T150_MC_BODEGAS.F150_ID = BI_T400.F_ID_BODEGA And " +
        '                "T150_MC_BODEGAS.F150_ROWID = T155_MC_UBICACION_AUXILIARES.F155_ROWID_BODEGA " +
        '            "Where " +
        '                "BI_T400.F_ID_CIA = 6 And " +
        '                "BI_T400.F_ID_BODEGA = 'B09' And " +
        '                "BI_T400.F_PARAMETRO_BIABLE = '6' And " +
        '                "BI_T400.F_ID_UBICACION_AUX Not Like '%CHATARRA%' And " +
        '                "BI_T400.F_CANT_DISPONIBLE_1 > 0 "

        'Consulta Resumida. No genera TimeOut
        Dim Slrs10 = "Select " +
                        "T120_MC_ITEMS.F120_DESCRIPCION, " +
                        "BI_T400.F_CANT_DISPONIBLE_1 " +
                    "From " +
                        "T120_MC_ITEMS Inner Join " +
                        "BI_T400 " +
                        "On BI_T400.F_ROWID_ITEM = T120_MC_ITEMS.F120_ROWID " +
                        "Where " +
                        "BI_T400.F_ID_CIA = 6 And " +
                        "BI_T400.F_ID_BODEGA = 'B09' And " +
                        "BI_T400.F_PARAMETRO_BIABLE = '6' And " +
                        "BI_T400.F_ID_UBICACION_AUX Not Like '%CHATARRA%' And " +
                        "BI_T400.F_CANT_DISPONIBLE_1 > 0 "
        'Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        Dim Str_Con_BdERP = CreaRQI.MyCommands.AbreDB_ERP

        'rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablaItems(,) As Object
        ReDim Preserve tablaItems(1, 0)
        Using connection As New SqlConnection(Str_Con_BdERP)
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
        Dim filename = CreaRQI.MyCommands.RutaPlanoAcad
        If filename = Nothing Then
            Exit Sub
        End If
        Dim extension As String
        If ADAPO = True Then
            extension = ".xls"
        Else
            extension = ".xls" '".txt"
        End If
        filename = filename & "Lista Stock " & extension

        Dim filenamesplit = Split(filename, ".")

        filename = filenamesplit(0)
        For i = 1 To UBound(filenamesplit) - 1
            filename = filename & "." & filenamesplit(i)
            'i = i + 1
        Next
        'revisa hasta 100 veces si el archivo existe
        For i = 0 To 100
            If My.Computer.FileSystem.FileExists(filename & ".xls") Then
                filename = filename & i
            Else
                Exit For
            End If
        Next
        filename = filename & ".xls"

        CreaRQI.MyCommands.escribe(filename, lista, True)
        If filename <> "" Then
            MsgBox("se ha guardado el listado con el nombre " & filename)
        End If
        Me.Close()
        If ADAPO = True Then
            Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication.activedocument.sendcommand("ADAPO" & vbCr)
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

        For i = 0 To CheckedListBox1.Items.Count - 1
            Dim raya = CheckedListBox1.Items.Item(i)
            If CheckedListBox1.GetItemChecked(i) And CheckedListBox1.Items.Item(i) < 51 Then
                Dim Cons As String = "Select ISNULL(Memo_Det.Memo_DetCantPFin, Saldos.cant) As Cant, Saldos.Tipo, Saldos.Anc1, Saldos.alto1, Saldos.Alto2, Saldos.Anc2, " +
                "Saldos.Plano_esp, IIF(saldos.Observacion = '.', '', Saldos.Observacion) AS Observacion, Saldos.Grupo, Saldos.Area, Orden.Ofa, Saldos.Item, Saldos.Ult_Memo, Memo_Det.Memo_DetOperacion, Memo_Det.Memo_Obs, Orden.Id_Of_P, Orden.Ref, Formaleta_Lib_Inv.Formaleta_Lib_Inv_Peso AS PesoUni, Formaleta_Lib_Inv.Formaleta_Lib_Inv_Peso * Saldos.Cant_Final_Req AS PesoTotal " +
                "FROM Explo_Saldos WITH (nolock) INNER JOIN Formaleta_Lib_Inv WITH (nolock) ON Explo_Saldos.Explo_Lib_Id = Formaleta_Lib_Inv.Formaleta_Lib_Inv_Id RIGHT OUTER JOIN " +
                    "Saldos WITH (nolock) INNER JOIN Orden WITH (nolock) ON Saldos.Id_Ofa = Orden.Id_Ofa ON Explo_Saldos.Saldos_Id = Saldos.Identificador LEFT OUTER JOIN " +
                    "Memos WITH (nolock) INNER JOIN Memo_Det WITH (nolock) ON Memos.Id_Memo = Memo_Det.Id_MemoId ON Saldos.Ult_Memo = Memos.Memo_No AND Saldos.Identificador = Memo_Det.Memo_DetIdSaldosOri " +
                "WHERE (Orden.Ofa LIKE '" & ComboBox2.SelectedItem & "-" & CheckedListBox1.Items.Item(i) & "') AND (Saldos.Anula = 0) OR " +
                    "(Orden.Ofa LIKE '" & ComboBox2.SelectedItem & "-" & CheckedListBox1.Items.Item(i) & "M') AND (Saldos.Anula = 0) " +
                "ORDER BY Saldos.Grupo, Saldos.Item"

                '"SELECT Saldos.cant, Saldos.Tipo, Saldos.Anc1, Saldos.alto1, Saldos.Alto2, Saldos.Anc2, Saldos.Observacion, Saldos.Grupo " + _
                '"FROM  Saldos INNER JOIN " + _
                '"Orden ON Saldos.Id_Ofa = Orden.Id_Ofa " + _
                '"WHERE (Orden.Ofa = '" & ComboBox2.SelectedItem & "-" & CheckedListBox1.Items.Item(i) & "')"

                Using connection As New SqlConnection(Str_Con_Bd)
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

        For i = 0 To CheckedListBox1.Items.Count - 1
            Dim raya = CheckedListBox1.Items.Item(i)
            If CheckedListBox1.GetItemChecked(i) And CheckedListBox1.Items.Item(i) > 50 Then
                Dim Cons As String = "SELECT DISTINCT ISNULL(Memo_Acc_Det.Memo_Acc_Cant_Fin, Of_Accesorios.Cant_Req) AS Cant_Req, CASE WHEN Of_Accesorios.Nomenclatura != '' THEN Of_Accesorios.Nomenclatura " +
                "WHEN Accesorios_Codigos.Nomenclatura !='' THEN Accesorios_Codigos.Nomenclatura ELSE item_planta.descripcion END AS Nomenclatura," +
                "Of_Accesorios.Des_Aux, Of_Accesorios.Dim1, Of_Accesorios.Dim2, Of_Accesorios.Dim3, Of_Accesorios.Dim4, Of_Accesorios.Dim5, " +
                "Of_Accesorios.Dim6, Orden.Ofa, Of_Accesorios.No_Item, Orden.Id_Of_P, Orden.Id_Ofa, Of_Accesorios.Ult_Memo, Memo_Acc_Det.Memo_Acc_Proceso," +
                "Memo_Acc_Det.Memo_Acc_Obs, Of_Accesorios.Peso_Estimado AS PesoUni, Of_Accesorios.Peso_Estimado * Of_Accesorios.Cant_Req AS PesoTotal " +
                "FROM  Of_Accesorios WITH (nolock) INNER JOIN " +
                "Orden WITH (nolock) ON Of_Accesorios.Id_Ofa = Orden.Id_Ofa INNER JOIN " +
                "item_planta WITH (nolock) ON Orden.planta_id = item_planta.planta_id AND Of_Accesorios.Id_UnoE = item_planta.cod_erp LEFT OUTER JOIN " +
                "Memo_Acc_Det WITH (nolock) ON Of_Accesorios.Id_Orden_Acce = Memo_Acc_Det.Memo_Acc_OfAccId LEFT OUTER JOIN " +
                "Accesorios_Codigos WITH (nolock) ON Of_Accesorios.Id_UnoE = Accesorios_Codigos.Id_UnoE AND Orden.planta_id = Accesorios_Codigos.planta_id " +
                "WHERE (Orden.Ofa LIKE '" & ComboBox2.SelectedItem & "-" & CheckedListBox1.Items.Item(i) & "%') AND (Of_Accesorios.Anula = 0) " +
                "ORDER BY Orden.Ofa, Of_Accesorios.No_Item"
                Using connection As New SqlConnection(Str_Con_Bd)
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
        If ComboBox1.SelectedItem = "ID" And idofaP = Nothing Then
            Dim cons = "SELECT DISTINCT Orden.Id_Of_P FROM Orden WITH (nolock) INNER JOIN Orden_Seg WITH (nolock) ON Orden.Id_Of_P = Orden_Seg.Id_Ofa WHERE (RTRIM(Orden_Seg.Num_Of) + '-' + RTRIM(Orden_Seg.Ano_Of) = '" & ComboBox2.SelectedItem & "')"
            Using connection As New SqlConnection(Str_Con_Bd)
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
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, Str_Con_Bd, "OG")
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, Str_Con_Bd, "OM")
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, Str_Con_Bd, "RC")
            CreaRQI.MyCommands.buscagarantiaalum(lista, idofaP, Str_Con_Bd, "ID")
            CreaRQI.MyCommands.buscagarantiaacc(lista, ComboBox2.SelectedItem, Str_Con_Bd)
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
                Exit Sub
            End If
            Dim extension As String
            If ADAPO = True Then
                extension = ".xls"
            Else
                extension = ".xls" '".txt"
            End If
            filename = filename & "Lista Orden " & ComboBox2.SelectedItem & extension
            Dim filenamesplit = Split(filename, ".")

            filename = filenamesplit(0)
            For i = 1 To UBound(filenamesplit) - 1
                filename = filename & "." & filenamesplit(i)
                'i = i + 1
            Next
            'revisa hasta 100 veces si el archivo existe
            For i = 0 To 100
                If My.Computer.FileSystem.FileExists(filename & ".xls") Then
                    filename = filename & i
                Else
                    Exit For
                End If
            Next
            filename = filename & ".xls"
            CreaRQI.MyCommands.escribe(filename, lista, True)
            If filename <> "" Then
                MsgBox("se ha guardado el listado con el nombre " & filename)
            End If
            Me.Close()
            If ADAPO = True Then
                Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication.activedocument.sendcommand("ADAPO" & vbCr)
            End If
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
        If CheckBox3.Checked = True Then
            descargalistaERP(True)
        Else
            descargalista(True)
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ComboBox1.Enabled = False
            ComboBox2.Enabled = False
            CheckBox1.Enabled = False
            CheckBox2.Enabled = False
            CheckedListBox1.Enabled = False
        Else
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            CheckBox1.Enabled = True
            CheckBox2.Enabled = True
            CheckedListBox1.Enabled = True
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        Carga_Checklist(ComboBox1.SelectedItem, ComboBox2.SelectedItem)
    End Sub

    Private Sub CheckedListBox1_Itemcheck(sender As Object, e As EventArgs) Handles CheckedListBox1.ItemCheck

    End Sub

    Private Sub Frm_ListFormaletas_Load(sender As Object, e As EventArgs) Handles Me.Load

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
        If CheckBox3.Checked = True Then
            descargalistaERP(False)
        Else
            descargalista(False)
        End If
    End Sub
End Class