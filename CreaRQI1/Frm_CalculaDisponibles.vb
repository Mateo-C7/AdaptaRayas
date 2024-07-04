Imports System.IO
Imports form1
Public Class Frm_CalculaDisponibles
    Dim bodegausada As String
    Private Sub Frm_CalculaDisponibles_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.CenterToScreen()
        'ComboBox1.Text = "Seleccione Compañía"
        CargaCias()
        For i = 0 To ComboBox1.Items.Count - 1
            Dim valor = Split(ComboBox1.Items.Item(i).ToString, " ")
            If valor(0) = Frm_CargaListados.CIA Then
                ComboBox1.SelectedIndex = i
                Exit For
            End If
        Next
    End Sub
    Sub CargaCias()
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
                "(T010_MM_COMPANIAS.F010_ID = 14) " + _
            "Order By " + _
                "T010_MM_COMPANIAS.F010_ID"
        Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablaCias(,) As Object
        tablaCias = rsiferp2.GetRows
        'Dim i As Integer = tablaCias.GetUpperBound(1)
        If tablaCias.GetUpperBound(1) >= 0 Then
            ComboBox1.Items.Clear()
            For i = 0 To tablaCias.GetUpperBound(1)
                ComboBox1.Items.Add(tablaCias(0, i) & " - " & tablaCias(1, i))

            Next
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'Dim a As Double = CheckedListBox1.Height
        'Dim b As Double = CheckedListBox1.Location.Y
        'Dim c As Double = PictureBox1.Height
        'Dim d As Double = PictureBox1.Location.Y
        'Dim f As Double = Me.Height
        Dim CiaVec() As String = Split(ComboBox1.Text, " ")
        Dim Cia As String = CiaVec(0)
        CargaBodega(Cia)
        Dim picturelocation As New System.Drawing.Point(PictureBox1.Location.X, CheckedListBox1.Location.Y + CheckedListBox1.Height + 6)
        PictureBox1.Location = picturelocation
    End Sub
    Sub CargaBodega(ByVal Cia As String)
        Dim rsiferp2 = New ADODB.Recordset
        Dim Slrs10 = "Select " + _
                        "T150_MC_BODEGAS.F150_ID, " + _
                        "T150_MC_BODEGAS.F150_DESCRIPCION " + _
                    "From " + _
                        "T150_MC_BODEGAS " + _
                    "Where " + _
                        "T150_MC_BODEGAS.F150_ID_CIA = " & Cia & " And " + _
                        "T150_MC_BODEGAS.F150_IND_MULTI_UBICACION = 0 And " + _
                        "T150_MC_BODEGAS.F150_IND_ESTADO = 1 " + _
                    "Order By " + _
                        "T150_MC_BODEGAS.F150_ID"
        Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
        rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim tablaCos(,) As Object
        tablaCos = rsiferp2.GetRows
        If tablaCos.GetUpperBound(1) >= 0 Then
            'ComboBox2.Items.Clear()
            CheckedListBox1.Items.Clear()

            For i = 0 To tablaCos.GetUpperBound(1)
                'ComboBox2.Items.Add(tablaCos(0, i) & " - " & tablaCos(1, i))
                CheckedListBox1.Items.Add(tablaCos(0, i) & " - " & tablaCos(1, i))
            Next

            If CheckedListBox1.Items.Count > 15 Then
                CheckedListBox1.ScrollAlwaysVisible = True
                CheckedListBox1.Height = 259

                MyBase.Size = New System.Drawing.Size(590, CheckedListBox1.Location.Y + CheckedListBox1.Height + PictureBox1.Height + 47)

            Else
                CheckedListBox1.ScrollAlwaysVisible = False
                CheckedListBox1.Height = (CheckedListBox1.Items.Count * 17) + 4

                'Dim tama As Integer
                'tama = CheckedListBox1.Height + 87

                'If tama < 160 Then
                'MyBase.Size = New System.Drawing.Size(590, 160)
                'Else
                'MyBase.Size = New System.Drawing.Size(590, tama)

                'End If
                MyBase.Size = New System.Drawing.Size(590, CheckedListBox1.Location.Y + CheckedListBox1.Height + PictureBox1.Height + 47)
            End If

        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CiaVec() As String = Split(ComboBox1.Text, " ")
        Dim Cia As String = CiaVec(0)
        Dim tablaItems As Object
        'Dim tablaItemsForline(,) As Object
        tablaItems = items(Cia)
        CreaRQI.MyCommands.verificaitems(False, False, Frm_CargaListados.grills, Frm_CargaListados.tabcontroluno, True, tablaItems, bodegausada, , Cia)
        Me.Close()
    End Sub
    Function items(ByVal Cia As String)
        Dim bodegausadavec()
        Dim tablaItemsFinal(,) As Object
        ReDim tablaItemsFinal(2, 0)
        bodegausada = ""
        For i = 0 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemChecked(i) Then
                If bodegausada = "" Then
                    bodegausadavec = Split(CheckedListBox1.Items.Item(i), " ")
                    For j = 1 To 5
                        If bodegausadavec(j) = "-" And bodegausadavec(j + 1) <> "" Then
                            bodegausada = bodegausadavec(j + 1)
                            Exit For
                        End If
                    Next

                End If
                Dim bodegavec() As String = Split(CheckedListBox1.Items.Item(i), " ")
                Dim bodega As String = bodegavec(0)

                Dim rsiferp2 = New ADODB.Recordset
                Dim Slrs10 = "Select " + _
                                "SALDOS_AL_DIA.NOMBRE_ITEM, " + _
                                "SALDOS_AL_DIA.EXISTENCIA, " + _
                                "SALDOS_AL_DIA.ITEM " + _
                            "From " + _
                                "SALDOS_AL_DIA " + _
                            "Where " + _
                                "SALDOS_AL_DIA.CIA = " & Cia & " And " + _
                                "SALDOS_AL_DIA.BODEGA = '" & bodega & "' And " + _
                                "SALDOS_AL_DIA.EXISTENCIA > 0 " + _
                            "Order By " + _
                                "SALDOS_AL_DIA.NOMBRE_ITEM"
                '"Select " + _
                '"SALDOS_ITEMS_ERP.NOMBRE_ITEM, " + _
                '"SALDOS_ITEMS_ERP.""Existencia_1"", " + _
                '"SALDOS_ITEMS_ERP.""Pendiente_1"" " + _
                '"From " + _
                '"SALDOS_ITEMS_ERP Inner Join " + _
                '"T150_MC_BODEGAS " + _
                '"On T150_MC_BODEGAS.F150_ID_CO = SALDOS_ITEMS_ERP.CO And " + _
                '"T150_MC_BODEGAS.F150_ID = SALDOS_ITEMS_ERP.BODEGA And " + _
                '"T150_MC_BODEGAS.F150_ID_CIA = SALDOS_ITEMS_ERP.EMPRESA " + _
                '"Where " + _
                '"SALDOS_ITEMS_ERP.EMPRESA = " & Cia & " And " + _
                '"SALDOS_ITEMS_ERP.TIPO_ITEM = 1 And " + _
                '"SALDOS_ITEMS_ERP.BODEGA = '" & bodega & "' And " + _
                '"SALDOS_ITEMS_ERP.F470_ID_PERIODO = 201606 And " + _
                '"SALDOS_ITEMS_ERP.F122_FACTOR <> 1 " + _
                '"Order By " + _
                '"SALDOS_ITEMS_ERP.NOMBRE_ITEM"
                Dim Con_Ora As ADODB.Connection = CreaRQI.MyCommands.AbreDb_Orl
                rsiferp2.Open(Slrs10, Con_Ora, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                Dim tablaItems(,) As Object
                tablaItems = rsiferp2.GetRows

                Dim sumado As Boolean

                For j = 0 To tablaItems.GetUpperBound(1)
                    sumado = False
                    For k = 0 To tablaItemsFinal.GetUpperBound(1)
                        If tablaItems(0, j) = tablaItemsFinal(1, k) Then
                            'Dim h As Integer = tablaItems(1, j) - tablaItems(2, j)
                            'If tablaItems(1, j) - tablaItems(2, j) <= 0 Then
                            'Exit For
                            'tablaItemsFinal(0, k) = 0
                            'Else
                            tablaItemsFinal(0, k) = tablaItems(1, j) + tablaItemsFinal(0, k) ' - tablaItems(2, j)
                            'End If

                            sumado = True
                        End If
                    Next
                    If sumado = False Then
                        'Dim h As Integer = tablaItems(1, j) - tablaItems(2, j)
                        'If tablaItems(1, j) - tablaItems(2, j) <= 0 Then
                        'Continue For
                        'tablaItemsFinal(0, tablaItemsFinal.GetUpperBound(1)) = 0
                        'Else
                        tablaItemsFinal(0, tablaItemsFinal.GetUpperBound(1)) = tablaItems(1, j) ' - tablaItems(2, j)
                        'End If
                        tablaItemsFinal(1, tablaItemsFinal.GetUpperBound(1)) = tablaItems(0, j)
                        tablaItemsFinal(2, tablaItemsFinal.GetUpperBound(1)) = tablaItems(2, j)
                        ReDim Preserve tablaItemsFinal(2, tablaItemsFinal.GetUpperBound(1) + 1)
                    End If
                Next

                Do While tablaItemsFinal(0, tablaItemsFinal.GetUpperBound(1)) = Nothing
                    ReDim Preserve tablaItemsFinal(2, tablaItemsFinal.GetUpperBound(1) - 1)
                Loop


            End If
        Next
        Return tablaItemsFinal
    End Function

    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As Windows.Forms.ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        If e.NewValue = Windows.Forms.CheckState.Checked Then
            For i = 0 To Me.CheckedListBox1.Items.Count - 1
                If i <> e.Index Then
                    Me.CheckedListBox1.SetItemChecked(i, False)
                End If
            Next
        End If
    End Sub
End Class