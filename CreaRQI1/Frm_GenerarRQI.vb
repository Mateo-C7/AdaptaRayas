Imports Autodesk.AutoCAD.ApplicationServices
Public Class Frm_GenerarRQI
    Dim terminar As Boolean
    Public Shared CIA As String
    Public Shared Id_Co As String
    Public Shared Id_Solicitante As String
    Public Shared Bodega_Salida As String
    Public Shared Bodega_Entrada As String
    Public Shared Cliente As String
    Public Shared Ubicacion As String
    Public Shared Referencia As String
    Public Shared notas As String
    Public Shared Fecha_Entrega As String
    Public Shared fecha As String = Format(Now(), "yyyyMMdd")
    Public Shared Num_Dias_Entrega As String
    

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles Me.Load
        
        DateTimePicker1.MinDate = Now().AddDays(1)
        DateTimePicker1.MaxDate = Now().AddDays(999)
        Me.CenterToScreen()
        Label1.Text = ""
        ComboBox2.Enabled = False
        ComboBox2.BackColor = Drawing.Color.Gray
        TextBox1.Enabled = False
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = False

        TextBox2.Enabled = False
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = False
        ComboBox3.BackColor = Drawing.Color.Gray
        ComboBox4.Enabled = False
        ComboBox4.BackColor = Drawing.Color.Gray
        ComboBox5.Enabled = False
        ComboBox5.BackColor = Drawing.Color.Gray
        ComboBox6.Enabled = False
        ComboBox6.BackColor = Drawing.Color.Gray
        TextBox3.Enabled = False
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = False
        ComboBox7.BackColor = Drawing.Color.Gray
        Button1.Enabled = False
        'terminar = False

        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " + _
                "T010_MM_COMPANIAS.F010_ID, " + _
                "T010_MM_COMPANIAS.F010_RAZON_SOCIAL " + _
            "From " + _
                "T010_MM_COMPANIAS " + _
            "Where " + _
                "(T010_MM_COMPANIAS.F010_ID = 15) Or " + _
                "(T010_MM_COMPANIAS.F010_ID = 11) Or " + _
                "(T010_MM_COMPANIAS.F010_ID = 12) Or " + _
                "(T010_MM_COMPANIAS.F010_ID = 14) Or " + _
                "(T010_MM_COMPANIAS.F010_ID = 6) " + _
            "Order By " + _
                "T010_MM_COMPANIAS.F010_ID"
        Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablacias As Object
        tablacias = rsiferp2.GetRows
        For i = 0 To UBound(tablacias, 2)
            ComboBox1.Items.Add(tablacias(0, i) & " - " & tablacias(1, i))
        Next
        For i = 0 To ComboBox1.Items.Count - 1
            Dim valor = Split(ComboBox1.Items.Item(i).ToString, " ")
            If valor(0) = Frm_CargaListados.CIA Then
                ComboBox1.SelectedIndex = i
                Exit For
            End If
        Next
        ComboBox7.Items.Add("ESTANDAR")
        ComboBox7.Items.Add("MUROS")
        ComboBox7.Items.Add("LOSAS")
    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = False
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = False
        TextBox2.Enabled = False
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = False
        ComboBox3.BackColor = Drawing.Color.Gray
        ComboBox4.Enabled = False
        ComboBox4.BackColor = Drawing.Color.Gray
        ComboBox5.Enabled = False
        ComboBox5.BackColor = Drawing.Color.Gray
        ComboBox6.Enabled = False
        ComboBox6.BackColor = Drawing.Color.Gray
        TextBox3.Enabled = False
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = False
        ComboBox7.BackColor = Drawing.Color.Gray
        Button1.Enabled = False
        CIA = selcia()
        llenacombo2(CIA)
    End Sub
    Sub llenacombo2(ByVal CIA As String)
        Dim CO As String
        ComboBox2.Items.Clear()
        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " + _
                    "T285_CO_CENTRO_OP.F285_ID, " + _
                    "T285_CO_CENTRO_OP.F285_DESCRIPCION " + _
                "From " + _
                    "T285_CO_CENTRO_OP " + _
                "Where " + _
                    "T285_CO_CENTRO_OP.F285_ID_CIA = " & CIA & " " + _
                "Order By " + _
                    "T285_CO_CENTRO_OP.F285_ID"
        Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablacos As Object
        Try
            tablacos = rsiferp2.GetRows
            If UBound(tablacos, 2) >= 0 Then
                For i = 0 To UBound(tablacos, 2)
                    ComboBox2.Items.Add(tablacos(0, i) & " - " & tablacos(1, i))
                Next
            Else
                CO = "000"
            End If
        Catch ex As Exception
            CO = "000"
        End Try

    End Sub
    Function selcia() As String
        Dim selciavec() As String
        Dim lentext As String
        selcia = ComboBox1.SelectedItem.ToString
        selciavec = Split(selcia, " ")
        lentext = selciavec(0).Length
        If lentext = 1 Then
            selcia = "00" & selciavec(0)
        ElseIf lentext = 2 Then
            selcia = "0" & selciavec(0)
        ElseIf lentext = 3 Then
            selcia = selciavec(0)
        End If
    End Function

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = True
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = False
        TextBox2.Enabled = False
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = False
        ComboBox3.BackColor = Drawing.Color.Gray
        ComboBox4.Enabled = False
        ComboBox4.BackColor = Drawing.Color.Gray
        ComboBox5.Enabled = False
        ComboBox5.BackColor = Drawing.Color.Gray
        ComboBox6.Enabled = False
        ComboBox6.BackColor = Drawing.Color.Gray
        TextBox3.Enabled = False
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = False
        ComboBox7.BackColor = Drawing.Color.Gray
        Button1.Enabled = False
        Label1.Text = ""
        ComboBox3.Items.Clear()
        Id_Co = SelCo(CIA)
        llenaBodega(True)
        TextBox1.Focus()
    End Sub
    Function SelCo(ByVal CIA As String) As String
        Dim selcovec() As String
        Dim lentext As String
        SelCo = ComboBox2.SelectedItem.ToString
        selcovec = Split(SelCo, " ")
        lentext = selcovec(0).Length
        If lentext = 1 Then
            SelCo = "00" & selcovec(0)
        ElseIf lentext = 2 Then
            SelCo = "0" & selcovec(0)
        ElseIf lentext = 3 Then
            SelCo = selcovec(0)
        End If
    End Function
    Sub llenaBodega(ByVal salida As Boolean)
        Dim salidaentrada As String
        If salida = True Then
            salidaentrada = "0"
        Else
            salidaentrada = "1"
        End If
        If CIA = 6 Then
            salidaentrada = 1
        End If
        Dim rsiferp2 = New ADODB.Recordset

        Dim Slrs10 = "Select " + _
                    "T150_MC_BODEGAS.F150_ID, " + _
                    "T150_MC_BODEGAS.F150_DESCRIPCION " + _
                "From " + _
                    "T150_MC_BODEGAS " + _
                "Where "
        If salidaentrada = False Then
            Slrs10 = Slrs10 & "T150_MC_BODEGAS.F150_ID_CO = " & Id_Co & " And "
        End If
        Slrs10 = Slrs10 & "T150_MC_BODEGAS.F150_ID_CIA = " & CIA & " And " + _
                    "T150_MC_BODEGAS.F150_IND_MULTI_UBICACION = " & salidaentrada & " And " + _
                    "T150_MC_BODEGAS.F150_IND_ESTADO = 1 " + _
                "Order by " + _
                    "T150_MC_BODEGAS.F150_ID_CIA, " + _
                    "T150_MC_BODEGAS.F150_ID_CO"
        Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablabodegas As Object
        Try
            tablabodegas = rsiferp2.GetRows
        Catch ex As Exception
            MsgBox("No existen Bodegas para el Centro de Operacion " & Id_Co, vbCritical, "Sin Bodegas")
            TextBox1.Enabled = False
            Exit Sub
        End Try
        If salida = True Then
            For i = 0 To UBound(tablabodegas, 2)
                ComboBox3.Items.Add(tablabodegas(0, i) & " - " & tablabodegas(1, i))
            Next
        Else
            ComboBox4.Items.Clear()
            For i = 0 To UBound(tablabodegas, 2)
                ComboBox4.Items.Add(tablabodegas(0, i) & " - " & tablabodegas(1, i))
            Next
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        Dim Id_Tercero As String
        If e.KeyValue = 13 Then
            Id_Tercero = SelTercero(CIA, Id_Co)
            If Id_Tercero = "NO" Then
                Label1.Text = "O"
                Label1.ForeColor = Drawing.Color.Red
                ComboBox2.Enabled = True
                ComboBox2.BackColor = Drawing.Color.White
                TextBox1.Enabled = True
                'TextBox1.BackColor = Drawing.Color.Gray
                DateTimePicker1.Enabled = False
                TextBox2.Enabled = False
                'TextBox2.BackColor = Drawing.Color.Gray
                ComboBox3.Enabled = False
                ComboBox3.BackColor = Drawing.Color.Gray
                ComboBox4.Enabled = False
                ComboBox4.BackColor = Drawing.Color.Gray
                ComboBox5.Enabled = False
                ComboBox5.BackColor = Drawing.Color.Gray
                ComboBox6.Enabled = False
                ComboBox6.BackColor = Drawing.Color.Gray
                TextBox3.Enabled = False
                'TextBox3.BackColor = Drawing.Color.Gray
                ComboBox7.Enabled = False
                ComboBox7.BackColor = Drawing.Color.Gray
                Button1.Enabled = False
                TextBox1.Focus()
            Else
                ComboBox2.Enabled = True
                ComboBox2.BackColor = Drawing.Color.White
                TextBox1.Enabled = True
                'TextBox1.BackColor = Drawing.Color.Gray
                DateTimePicker1.Enabled = True
                DateTimePicker1.MinDate = Now().AddDays(1)
                TextBox2.Enabled = True
                'TextBox2.BackColor = Drawing.Color.Gray
                ComboBox3.Enabled = False
                ComboBox3.BackColor = Drawing.Color.Gray
                ComboBox4.Enabled = False
                ComboBox4.BackColor = Drawing.Color.Gray
                ComboBox5.Enabled = False
                ComboBox5.BackColor = Drawing.Color.Gray
                ComboBox6.Enabled = False
                ComboBox6.BackColor = Drawing.Color.Gray
                TextBox3.Enabled = False
                'TextBox3.BackColor = Drawing.Color.Gray
                ComboBox7.Enabled = False
                ComboBox7.BackColor = Drawing.Color.Gray
                Button1.Enabled = False
                Label1.Text = "P"
                Label1.ForeColor = Drawing.Color.Green
                Id_Solicitante = SelTercero(CIA, Id_Co, Id_Tercero)
            End If
        ElseIf e.KeyCode = Windows.Forms.Keys.Escape Then
            Me.Close()
        End If
    End Sub
    Function SelTercero(ByVal CIA As String, CO As String, Optional ByVal Id_Tercero As String = "") As String
        Dim existeCedula As Boolean
        existeCedula = False
        Dim tercero As String
        Dim terceroAux() As String
        Dim terceroVector1(14) As String
        Dim terceroVector2() As String
        For i = 0 To 14
            terceroVector1(i) = " "
        Next
        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " + _
                    "T200_MM_TERCEROS.F200_ID, " + _
                    "T211_MM_FUNCIONARIOS.F211_ID " + _
                "From " + _
                    "T200_MM_TERCEROS Inner Join " + _
                    "T211_MM_FUNCIONARIOS " + _
                        "On T211_MM_FUNCIONARIOS.F211_ROWID_TERCERO = T200_MM_TERCEROS.F200_ROWID " + _
                "Where " + _
                    "T200_MM_TERCEROS.F200_ID_CIA = " & CIA & " And " + _
                    "T211_MM_FUNCIONARIOS.F211_IND_SOLICITANTE = 1"
        Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablaterceros As Object
        tablaterceros = rsiferp2.GetRows
        If Id_Tercero = "" Then
            tercero = TextBox1.Text
            For i = 0 To UBound(tablaterceros, 2)
                terceroAux = Split(tablaterceros(0, i), " ")
                If tercero = terceroAux(0) Then
                    existeCedula = True
                    Exit For
                End If
            Next
            If existeCedula Then
                terceroVector2 = CreaRQI.MyCommands.SeparaCaracteres(tercero)


                Dim j = Len(tercero) - 1
                'k = 14
                SelTercero = ""
                For i = 0 To j
                    terceroVector1(i) = terceroVector2(i)
                Next

                For j = 0 To UBound(terceroVector1)
                    SelTercero = SelTercero & terceroVector1(j)
                Next
            Else
                MsgBox("Digite una Cédula Válida", vbCritical, "Número de Cédula no Existe")
                TextBox1.Text = ""
                SelTercero = "NO"
                Exit Function
            End If
        Else
            terceroVector2 = Split(Id_Tercero, " ")
            For i = 0 To UBound(tablaterceros, 2)
                If InStr(tablaterceros(0, i), terceroVector2(0)) Then
                    terceroAux = CreaRQI.MyCommands.SeparaCaracteres(tablaterceros(1, i))
                    SelTercero = ""
                    For j = 0 To UBound(terceroAux) '- 1
                        If terceroAux(j) <> " " Then
                            SelTercero = SelTercero & terceroAux(j)
                        End If
                    Next
                    If Len(SelTercero) = 1 Then
                        SelTercero = SelTercero & "    "
                    ElseIf Len(SelTercero) = 2 Then
                        SelTercero = SelTercero & "   "
                    ElseIf Len(SelTercero) = 3 Then
                        SelTercero = SelTercero & "  "
                    ElseIf Len(SelTercero) = 4 Then
                        SelTercero = SelTercero & " "
                    End If
                    Exit For
                End If
            Next
        End If
    End Function


    Private Sub TextBox2_DoubleClick(sender As Object, e As EventArgs) Handles TextBox2.DoubleClick

    End Sub



    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = True
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = True
        TextBox2.Enabled = True
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = True
        ComboBox3.BackColor = Drawing.Color.White
        ComboBox4.Enabled = False
        ComboBox4.BackColor = Drawing.Color.Gray
        ComboBox5.Enabled = False
        ComboBox5.BackColor = Drawing.Color.Gray
        ComboBox6.Enabled = False
        ComboBox6.BackColor = Drawing.Color.Gray
        TextBox3.Enabled = False
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = False
        ComboBox7.BackColor = Drawing.Color.Gray
        Button1.Enabled = False

        ComboBox3.Focus()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = True
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = True
        TextBox2.Enabled = True
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = True
        ComboBox3.BackColor = Drawing.Color.White
        ComboBox4.Enabled = True
        ComboBox4.BackColor = Drawing.Color.White
        ComboBox5.Enabled = False
        ComboBox5.BackColor = Drawing.Color.Gray
        ComboBox6.Enabled = False
        ComboBox6.BackColor = Drawing.Color.Gray
        TextBox3.Enabled = False
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = False
        ComboBox7.BackColor = Drawing.Color.Gray
        Button1.Enabled = False
        Fecha_Entrega = Format(DateTimePicker1.Value, "yyyyMMdd")
        Dias_Entrega()
        Bodega_Salida = SelBodega(True)
        llenaBodega(False)
    End Sub
    Function SelBodega(ByVal salida As Boolean) As String
        Dim selBodegaVec() As String
        Dim lentext As Integer
        If salida = True Then
            SelBodega = ComboBox3.SelectedItem.ToString
        Else
            SelBodega = ComboBox4.SelectedItem.ToString
        End If
        selBodegaVec = Split(SelBodega, " ")
        lentext = selBodegaVec(0).Length
        If lentext = 1 Then
            SelBodega = selBodegaVec(0) & "    "
        ElseIf lentext = 2 Then
            SelBodega = selBodegaVec(0) & "   "
        ElseIf lentext = 3 Then
            SelBodega = selBodegaVec(0) & "  "
        ElseIf lentext = 4 Then
            SelBodega = selBodegaVec(0) & " "
        ElseIf lentext = 5 Then
            SelBodega = selBodegaVec(0)
        End If
    End Function

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = True
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = True
        TextBox2.Enabled = True
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = True
        ComboBox3.BackColor = Drawing.Color.White
        ComboBox4.Enabled = True
        ComboBox4.BackColor = Drawing.Color.White
        ComboBox5.Enabled = True
        ComboBox5.BackColor = Drawing.Color.White
        ComboBox6.Enabled = False
        ComboBox6.BackColor = Drawing.Color.Gray
        TextBox3.Enabled = False
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = False
        ComboBox7.BackColor = Drawing.Color.Gray
        Button1.Enabled = False
        Bodega_Entrada = SelBodega(False)
        LlenaUbica()
    End Sub
    Sub LlenaUbica(Optional ByVal Cliente As String = "")
        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " + _
                    "T155_MC_UBICACION_AUXILIARES.F155_ID, " + _
                    "T155_MC_UBICACION_AUXILIARES.F155_DESCRIPCION, " + _
                    "T201_MM_CLIENTES.F201_DESCRIPCION_SUCURSAL, " + _
                    "T201_MM_CLIENTES.F201_ROWID_TERCERO " + _
                "From " + _
                    "T155_MC_UBICACION_AUXILIARES Inner Join " + _
                    "T150_MC_BODEGAS " + _
                    "On T155_MC_UBICACION_AUXILIARES.F155_ROWID_BODEGA = " + _
                    "T150_MC_BODEGAS.F150_ROWID Inner Join " + _
                    "T201_MM_CLIENTES " + _
                    "On T155_MC_UBICACION_AUXILIARES.F155_ROWID_TERCERO = " + _
                    "T201_MM_CLIENTES.F201_ROWID_TERCERO And " + _
                    "T155_MC_UBICACION_AUXILIARES.F155_ID_SUCURSAL = " + _
                    "T201_MM_CLIENTES.F201_ID_SUCURSAL " + _
                "Where " + _
                    "T155_MC_UBICACION_AUXILIARES.F155_ID_CIA = " & CIA & " And " + _
                    "T150_MC_BODEGAS.F150_IND_MULTI_UBICACION = 1 And " + _
                    "T150_MC_BODEGAS.F150_ID = '" & Bodega_Entrada & "' " + _
                "Order by " + _
                    "T201_MM_CLIENTES.F201_DESCRIPCION_SUCURSAL"
        Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim TablaUbica As Object
        Dim TablaClientes As Object
        ReDim TablaClientes(0)
        TablaUbica = rsiferp2.GetRows
        TablaClientes(0) = TablaUbica(2, 0) & " " & TablaUbica(3, 0)
        If Cliente = "" Then
            TablaClientes(0) = TablaUbica(2, 0) & " " & TablaUbica(3, 0)
            For i = 0 To UBound(TablaUbica, 2) - 1
                If TablaClientes(UBound(TablaClientes)) = TablaUbica(2, i + 1) & " " & TablaUbica(3, i + 1) Then
                Else
                    ReDim Preserve TablaClientes(UBound(TablaClientes) + 1)
                    TablaClientes(UBound(TablaClientes)) = TablaUbica(2, i + 1) & " " & TablaUbica(3, i + 1)
                End If
            Next
            For i = 0 To UBound(TablaClientes)
                ComboBox5.Items.Add(TablaClientes(i))
            Next
        Else
            For i = 0 To UBound(TablaUbica, 2)
                If TablaUbica(3, i) = Cliente Then
                    TablaClientes(UBound(TablaClientes)) = TablaUbica(0, i) & "- " & TablaUbica(1, i)
                    ReDim Preserve TablaClientes(UBound(TablaClientes) + 1)
                End If
            Next
            For i = 0 To UBound(TablaClientes)
                Try
                    ComboBox6.Items.Add(TablaClientes(i))
                Catch ex As Exception

                End Try
            Next
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = True
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = True
        TextBox2.Enabled = True
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = True
        ComboBox3.BackColor = Drawing.Color.White
        ComboBox4.Enabled = True
        ComboBox4.BackColor = Drawing.Color.White
        ComboBox5.Enabled = True
        ComboBox5.BackColor = Drawing.Color.White
        ComboBox6.Enabled = True
        ComboBox6.BackColor = Drawing.Color.White
        TextBox3.Enabled = False
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = False
        ComboBox7.BackColor = Drawing.Color.Gray
        Button1.Enabled = False
        Cliente = SelCliente()
        LlenaUbica(Cliente)
    End Sub
    Function SelCliente() As String
        Dim clienteVector As Object
        clienteVector = Split(ComboBox5.SelectedItem.ToString, " ")
        SelCliente = clienteVector(UBound(clienteVector))
    End Function

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = True
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = True
        TextBox2.Enabled = True
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = True
        ComboBox3.BackColor = Drawing.Color.White
        ComboBox4.Enabled = True
        ComboBox4.BackColor = Drawing.Color.White
        ComboBox5.Enabled = True
        ComboBox5.BackColor = Drawing.Color.White
        ComboBox6.Enabled = True
        ComboBox6.BackColor = Drawing.Color.White
        TextBox3.Enabled = True
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = True
        ComboBox7.BackColor = Drawing.Color.White
        Button1.Enabled = False
        Ubicacion = SelUbica()
    End Sub
    Function SelUbica() As String
        Dim UbicaVector() = Split(ComboBox6.SelectedItem.ToString, " ")
        Dim ubiVector1(9) As String
        For i = 0 To 9
            ubiVector1(i) = " "
        Next
        Dim ubiVector2() As String = CreaRQI.MyCommands.SeparaCaracteres(UbicaVector(0))
        Dim j = Len(UbicaVector(0)) - 1
        'k = 9
        SelUbica = ""
        For i = 0 To j
            ubiVector1(i) = ubiVector2(i)
            'j = j - 1
            'k = k - 1
        Next
        For i = 0 To UBound(ubiVector1)
            SelUbica = SelUbica & ubiVector1(i)
        Next
    End Function

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ComboBox2.Enabled = True
        ComboBox2.BackColor = Drawing.Color.White
        TextBox1.Enabled = True
        'TextBox1.BackColor = Drawing.Color.Gray
        DateTimePicker1.Enabled = True
        TextBox2.Enabled = True
        'TextBox2.BackColor = Drawing.Color.Gray
        ComboBox3.Enabled = True
        ComboBox3.BackColor = Drawing.Color.White
        ComboBox4.Enabled = True
        ComboBox4.BackColor = Drawing.Color.White
        ComboBox5.Enabled = True
        ComboBox5.BackColor = Drawing.Color.White
        ComboBox6.Enabled = True
        ComboBox6.BackColor = Drawing.Color.White
        TextBox3.Enabled = True
        'TextBox3.BackColor = Drawing.Color.Gray
        ComboBox7.Enabled = True
        ComboBox7.BackColor = Drawing.Color.White
        Button1.Enabled = True
        Referencia = selRef()
    End Sub
    Function selRef() As String
        Dim RefVector1(19) As String
        For i = 0 To 19
            RefVector1(i) = " "
        Next
        Dim RefVector2() As String = CreaRQI.MyCommands.SeparaCaracteres(ComboBox7.SelectedItem.ToString)
        Dim j = Len(ComboBox7.SelectedItem.ToString) - 1
        'k = 19
        selRef = ""
        For i = 0 To j
            RefVector1(i) = RefVector2(i)
            'j = j - 1
            'k = k - 1
        Next
        For i = 0 To UBound(RefVector1)
            selRef = selRef & RefVector1(i)
        Next
    End Function
    Public Shared cargada As Boolean
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        notas = SelNotas()
        CreaRQI.MyCommands.CreaPlano()
        If cargada = True Then
            Me.Close()
        End If
    End Sub
    Function SelNotas() As String
        Dim NotasVector1(254) As String
        For i = 0 To 254
            NotasVector1(i) = " "
        Next
        Dim NotasVector2() As String
        NotasVector2 = CreaRQI.MyCommands.SeparaCaracteres(TextBox3.Text)
        Dim j = TextBox3.Text.Length - 1
        'k = 254
        SelNotas = ""
        For i = 0 To j
            NotasVector1(i) = NotasVector2(i)
            'j = j - 1
            'k = k - 1
        Next
        For i = 0 To UBound(NotasVector1)
            SelNotas = SelNotas & NotasVector1(i)
        Next
    End Function
    Public forma2 As System.Windows.Forms.Form

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Dim valor = MsgBox("¿desea salir de la generacion de RQI?", MsgBoxStyle.OkCancel)
        If valor = Microsoft.VisualBasic.MsgBoxResult.Cancel Then
            e.Cancel = True
        Else
            forma2.Enabled = True
            'habilitar form
        End If
    End Sub
    Sub mensajesalida(ByVal e As Windows.Forms.KeyEventArgs)
        Try
            If e.KeyCode = Windows.Forms.Keys.Escape Then
                Me.Close()
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Form3_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles ComboBox2.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs)
        mensajesalida(e)
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub ComboBox3_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles ComboBox3.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub ComboBox4_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles ComboBox4.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub ComboBox5_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles ComboBox5.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub ComboBox6_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles ComboBox6.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub ComboBox7_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles ComboBox7.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub Button1_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles Button1.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub Button2_KeyDown(sender As Object, e As Windows.Forms.KeyEventArgs) Handles Button2.KeyDown
        mensajesalida(e)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    'Private Sub Form3_FormClosing(sender As Object, e As Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
    'Me.Close()
    'End Sub
    Sub Dias_Entrega()
        Dim bidias As Date
        Dim sumardias = 0
        If DateTimePicker1.Value.Year > Now().Year Then
            For i = Now().Year To DateTimePicker1.Value.Year - 1
                Try
                    bidias = CDate("29-02-" & CStr(i))
                    sumardias = sumardias + 366
                Catch ex As Exception
                    sumardias = sumardias + 365
                End Try
            Next
        End If
        Dim Dias_Entrega = DateTimePicker1.Value.DayOfYear + sumardias - Now().DayOfYear
        If Dias_Entrega.ToString.Length = 1 Then
            Num_Dias_Entrega = "  " & Dias_Entrega.ToString
        ElseIf Dias_Entrega.ToString.Length = 2 Then
            Num_Dias_Entrega = " " & Dias_Entrega.ToString
        Else
            Num_Dias_Entrega = Dias_Entrega.ToString
        End If
    End Sub

End Class